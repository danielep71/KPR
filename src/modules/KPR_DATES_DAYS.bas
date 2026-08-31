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
'     callers and legacy single-cell references keep working. This applies to
'     every Variant argument, the pillar token included.
'   - Multi-cell / multi-element inputs are rejected as #VALUE!.
'   - Array traversal is deliberately deferred to issue #16. When it arrives it
'     wraps these scalar cores rather than duplicating them.
'
' DESIGN / INPUT NORMALIZATION
'   - Every DateIn-style argument is a Variant funnelled through TryResolveDate,
'     which composes the three boundaries in a fixed order:
'         unwrap shape -> propagate incoming error -> strict parse/window gate
'   - KPR_Core_Parse applies the window once after date-only normalization. Its
'     serial/year constants are statically pinned to the Date constants owned by
'     KPR_Core_Dates because the dependency matrix forbids a direct call.
'   - The two weekday locators take a year and a month rather than a date, so
'     they never reach TryResolveDate. They gate their own result against the
'     window instead. Without that gate a year check alone is too coarse: the
'     window floor is 1 March 1900, so January and February 1900 pass a
'     year-granularity test and would return an out-of-window date.
'
' ERROR POLICY (USER FACING)
'   Public UDFs return either a valid scalar result or a native Excel error.
'   No public function returns a message string.
'
'       #VALUE!  the input cannot be interpreted
'                (unparseable date, non-scalar input, malformed pillar, month
'                 outside 1..12, weekday outside 1..7, occurrence outside 1..5,
'                 unexpected runtime error)
'
'       #NUM!    the input is well formed but the answer does not exist or
'                falls outside the supported date window
'                (no such weekday occurrence in the month, a located weekday
'                 outside the window, arithmetic or pillar shift beyond
'                 KPR_MIN_DATE / KPR_MAX_DATE)
'
'   Rationale: message strings are invisible to IFERROR / ISERROR, break any
'   downstream arithmetic that references the cell, and cannot be aggregated or
'   filtered. Native error values are the only return that composes.
'
' CALENDAR ARITHMETIC
'   No function in this module writes the day-0 idiom DateSerial(y, m + 1, 0).
'   It raises error 5 at both ends of the VBA Date range, and behind a
'   defensive handler that surfaces as a worksheet error indistinguishable
'   from a genuine rejection. Month length comes from KPR_Core_Dates.DaysInMonth
'   and month-end from KPR_Core_Dates.EndOfMonth, both of which are total.
'
' KNOWN DIVERGENCES FROM THE FROZEN CONTRACT
'   Issue #12 removed the input-handling divergences. What remains is owned by
'   the issues named against each line:
'
'     - Existing Boolean controls implement their strict Variant semantics;
'       Opt_Rounding and its token parser arrive with issue #14.
'     - Duplicate pillar units accumulate rather than being rejected.    (#15)
'     - Pillar rounding is NEAREST only; Opt_Rounding is absent.         (#14)
'     - No host date-system detection exists, so a 1904 workbook is answered
'       with silently shifted values.                                    (#17)
'     - Multi-cell inputs are rejected rather than traversed, and a control
'       larger than 1x1 reports a shape rejection rather than the contract's
'       distinct CONTROL_NOT_SCALAR condition.                           (#16)
'     - DaysInYear and IsLeapYear still take a date rather than a calendar
'       year, and DatesFromPillar keeps its plural name.                 (#15)
'
' INTERNAL ERROR POLICY
'   - Core routines are Try-style (Boolean return, result ByRef) or plain
'     computation on already-validated inputs.
'   - This facade owns all worksheet-facing error behavior.
'   - Defensive catch-all handlers are containment only. They must be
'     unreachable in contract-conforming execution; an activation is a defect,
'     never an expected outcome. A raise that a handler converts into a
'     plausible worksheet error is invisible from the sheet, so a condition
'     that can be tested is tested rather than trapped.
'
' DEPENDENCIES / INTEGRATION
'   - KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates, KPR_Core_Err
'   - Private helpers in this module: TryResolveDate, TryResolveLong,
'     TryResolveBool
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
    Optional ByVal Opt_WeekBaseMonday As Variant = True) _
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
' WHY THIS EXISTS
'   A sheet can reach VBA's Weekday only through WEEKDAY, whose return_type
'   codes are a different vocabulary from this library's. Exposing the choice
'   as one Boolean keeps the two conventions in the layer's own terms.
'
' SIGNATURE
'   KPR_Dates_DayOfWeek(DateIn, [Opt_WeekBaseMonday]) -> Variant
'
' INPUTS
'   DateIn
'     Date candidate handled by the facade TryResolveDate policy.
'
'   Opt_WeekBaseMonday (optional, default TRUE)
'     Weekday() base selector:
'       TRUE  => vbMonday (Mon=1..Sun=7)
'       FALSE => vbSunday (Sun=1..Sat=7)
'
' RETURNS
'   Variant
'     Long (1..7) on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar, or Opt_WeekBaseMonday is not an
'            omitted, Empty or native Boolean value
'   #NUM!    DateIn is outside the supported window
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Err.ErrValue
'
' NOTES
'   - TRUE selects Monday-based numbering (ISO habit); FALSE selects
'     Sunday-based numbering (US habit).
'   - The Variant declaration preserves the incoming runtime type. Numeric 0/1,
'     Boolean-looking text, dates and objects are rejected rather than coerced.
'   - Omission applies the declared default TRUE. Empty, including a reference
'     to a blank cell, selects the same documented default in TryResolveBool.
'   - Weekday is total for any Date once the base is one of the vbDayOfWeek
'     constants, so the handler below covers only failures inside
'     TryResolveDate.
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
    Dim WkMonday        As Boolean   'Strictly parsed Opt_WeekBaseMonday
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

    'Accept only an omitted, blank or native Boolean Opt_WeekBaseMonday
        If Not TryResolveBool(Opt_WeekBaseMonday, True, WkMonday, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' RESOLVE WEEKDAY BASE
'------------------------------------------------------------------------------
    'Resolve Weekday() base once
        If WkMonday Then
            WkBase = vbMonday
        Else
            WkBase = vbSunday
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return weekday index 1..7 under the selected base
        KPR_Dates_DayOfWeek = CLng(Weekday(ParsedDate, WkBase))
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
' WHY THIS EXISTS
'   Month length is the one calendar quantity a worksheet cannot derive without
'   restating the leap rule. Exposing it here means the sheet asks the same
'   DaysInMonth the shift layer uses, rather than a parallel formula.
'
' SIGNATURE
'   KPR_Dates_DaysInMonth(DateIn) -> Variant
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
'   #VALUE!  DateIn not parseable / not scalar, or unexpected runtime error
'   #NUM!    DateIn is outside the supported window
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.DaysInMonth
'   - KPR_Core_Err.ErrValue
'
' NOTES
'   - The core DaysInMonth is called directly rather than through EndOfMonth.
'     Reading the day number of a constructed month-end would build a date only
'     to take it apart again, and EndOfMonth derives that date from this same
'     function; the round trip adds a DateSerial and a Day per cell for no
'     information. No CLng is needed because the core member returns Long.
'   - The month passed to the core function comes from a parsed Date, so it is
'     always 1 to 12 and the invalid-month return of zero is unreachable here.
'   - Parse failures map to #VALUE! while DATE_WINDOW maps to #NUM!, preserving
'     the contract's distinction at the worksheet boundary.
'   - Fail and Err_Handler are kept separate so a contract-level rejection and
'     an unexpected raise cannot be confused while reading the flow. The
'     handler reasserts FailErr because a raise inside ErrValue itself would
'     otherwise leave it Empty.
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Month length straight from the calendar core; no date is constructed
        KPR_Dates_DaysInMonth = DaysInMonth(Year(ParsedDate), Month(ParsedDate))
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
' WHY THIS EXISTS
'   Year length is a day-count denominator, and a sheet deriving it inline is a
'   sheet restating the leap rule. Routing it through the core predicate means
'   this surface and the calendar layer can never disagree about a year.
'
' SIGNATURE
'   KPR_Dates_DaysInYear(DateIn) -> Variant
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
'   #VALUE!  DateIn not parseable / not scalar, or unexpected runtime error
'   #NUM!    DateIn is outside the supported window
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.IsLeapYear
'   - KPR_Core_Err.ErrValue
'
' NOTES
'   - The year is read from the parsed Date and passed as a calendar year, not
'     as a serial. IsLeapYear takes a year and would silently treat a serial as
'     one, so the Year() call is load-bearing rather than incidental.
'   - 1900 returns 365. It is divisible by 100 but not by 400, so the century
'     exception applies. Every year reachable through this surface is 1900 or
'     later, and the window floor of 1 March 1900 means year 1900 is reachable
'     even though February 1900 is not.
'   - IsLeapYear applies the rule proleptically, so it reports leap years for
'     dates before the 1582 Gregorian reform that the historical calendar does
'     not. No such year is reachable here; the caveat belongs to the core
'     predicate, not to this surface.
'   - Parse failures map to #VALUE! while DATE_WINDOW maps to #NUM!, preserving
'     the contract's distinction at the worksheet boundary.
'   - Fail and Err_Handler are kept separate so a contract-level rejection and
'     an unexpected raise cannot be confused while reading the flow. The
'     handler reasserts FailErr because a raise inside ErrValue itself would
'     otherwise leave it Empty.
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return 366 for leap years, 365 otherwise
        If IsLeapYear(Year(ParsedDate)) Then
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
' SIGNATURE
'   KPR_Dates_BeginOfMonth(DateIn) -> Variant
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
'   #VALUE!  DateIn not parseable / not scalar, or unexpected runtime error
'   #NUM!    DateIn is outside the supported window
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Err.ErrValue
'
' NOTES
'   - Only the year / month components of the parsed input survive; the result
'     is always day 1.
'   - The result can precede the window floor: any accepted date in March 1900
'     resolves to 1 March 1900, which is the floor itself, so no earlier value
'     is reachable. No result gate is required here.
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

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
' SIGNATURE
'   KPR_Dates_EndOfMonth(DateIn) -> Variant
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
'   #VALUE!  DateIn not parseable / not scalar, or unexpected runtime error
'   #NUM!    DateIn is outside the supported window
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.EndOfMonth
'   - KPR_Core_Err.ErrValue
'
' NOTES
'   - The core EndOfMonth is total across the whole VBA Date range, December
'     9999 included, so the top of the window needs no special handling here.
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the last day of the containing month
        KPR_Dates_EndOfMonth = EndOfMonth(ParsedDate)
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
' SIGNATURE
'   KPR_Dates_IsMonthEnd(DateIn) -> Variant
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
'   #VALUE!  DateIn not parseable / not scalar, or unexpected runtime error
'   #NUM!    DateIn is outside the supported window
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.DaysInMonth
'   - KPR_Core_Err.ErrValue
'
' NOTES
'   - The test compares day numbers rather than dates. Building the month-end
'     date only to compare it against the input costs a DateSerial per cell and
'     answers the same question.
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Compare the day number to the length of its own month
        KPR_Dates_IsMonthEnd = _
            (Day(ParsedDate) = DaysInMonth(Year(ParsedDate), Month(ParsedDate)))
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
' SIGNATURE
'   KPR_Dates_IsQuarterEnd(DateIn) -> Variant
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
'   #VALUE!  DateIn not parseable / not scalar, or unexpected runtime error
'   #NUM!    DateIn is outside the supported window
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.DaysInMonth
'   - KPR_Core_Err.ErrValue
'
' NOTES
'   - Quarter months are detected as Month Mod 3 = 0, which is exactly the set
'     {3, 6, 9, 12}. This is a calendar quarter, not a fiscal quarter; a fiscal
'     variant belongs in a separate function with an explicit anchor argument.
'   - The month test is nested outside the month-length lookup rather than
'     combined with And. VBA does not short-circuit And, so a single expression
'     would compute the month length for every input; nesting skips it for the
'     eleven months in twelve that cannot be a quarter end.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim MonthPart       As Long      'Calendar month of the parsed date
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Cheap test first; only a quarter month can be a quarter end
        MonthPart = Month(ParsedDate)
        If (MonthPart Mod 3) <> 0 Then
            KPR_Dates_IsQuarterEnd = False
        Else
            'Quarter month, so the answer is whether it is also month-end
                KPR_Dates_IsQuarterEnd = _
                    (Day(ParsedDate) = DaysInMonth(Year(ParsedDate), MonthPart))
        End If
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
' SIGNATURE
'   KPR_Dates_IsYearEnd(DateIn) -> Variant
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
'   #VALUE!  DateIn not parseable / not scalar, or unexpected runtime error
'   #NUM!    DateIn is outside the supported window
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Err.ErrValue
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

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
' SIGNATURE
'   KPR_Dates_IsLeapYear(DateIn) -> Variant
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
'   #VALUE!  DateIn not parseable / not scalar, or unexpected runtime error
'   #NUM!    DateIn is outside the supported window
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.IsLeapYear
'   - KPR_Core_Err.ErrValue
'
' NOTES
'   - The previous per-function year gate is gone. TryResolveDate already
'     refuses anything before KPR_MIN_DATE, so no year reaching this point can
'     be outside the supported range.
'   - The year is passed as a calendar year, not as a serial. IsLeapYear takes
'     a Long and would silently treat a serial as a year, so the Year() call is
'     load-bearing rather than incidental.
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Apply the Gregorian leap-year rule to the containing year
        KPR_Dates_IsLeapYear = IsLeapYear(Year(ParsedDate))
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
    ByVal nWeeks As Variant) _
    As Variant
'
'==============================================================================
' KPR_Dates_AddWeeks
'------------------------------------------------------------------------------
' PURPOSE
'   Adds nWeeks to DateIn using calendar-day arithmetic:
'       Result = DateIn + (7 * nWeeks) days
'
' SIGNATURE
'   KPR_Dates_AddWeeks(DateIn, nWeeks) -> Variant
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
'   #VALUE!  DateIn not parseable / not scalar, or unexpected runtime error
'   #NUM!    DateIn is outside the supported window
'   #NUM!    result falls outside KPR_MIN_DATE .. KPR_MAX_DATE
'
'   nWeeks is a Variant parsed strictly before conversion. Fractions and
'   Boolean values return #VALUE!; values outside the Long range return #NUM!.
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Err.ErrValue, KPR_Core_Err.ErrNum
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
    Dim nWeeksL         As Long      'Strictly parsed nWeeks
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

    'Reject a fractional, out-of-range or non-numeric nWeeks
        If Not TryResolveLong(nWeeks, nWeeksL, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' COMPUTE SHIFT
'------------------------------------------------------------------------------
    'Compute the shifted serial in Double to avoid overflow on coercion
        ResultD = CDbl(ParsedDate) + (7# * CDbl(nWeeksL))

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
    ByVal nMonths As Variant, _
    Optional ByVal Opt_KeepEOM As Variant = False) _
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
' SIGNATURE
'   KPR_Dates_AddMonths(DateIn, nMonths, [Opt_KeepEOM]) -> Variant
'
' INPUTS
'   DateIn
'     Date candidate handled by the facade TryResolveDate policy.
'
'   nMonths
'     Number of calendar months to add. Negative values are allowed.
'
'   Opt_KeepEOM (optional, default FALSE)
'     TRUE  => EOM in, EOM out
'     FALSE => clip day-of-month when the target month is shorter
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar, or unexpected runtime error
'   #NUM!    DateIn is outside the supported window
'   #NUM!    result falls outside KPR_MIN_DATE .. KPR_MAX_DATE
'
'   nMonths is a Variant parsed strictly before conversion. Fractions and
'   Boolean values return #VALUE!; values outside the Long range return #NUM!.
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.TryAddMonths
'   - KPR_Core_Err.ErrValue, KPR_Core_Err.ErrNum
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
'   - Omitting Opt_KeepEOM and referencing a blank control cell both select the
'     declared False default. No Boolean coercion is performed.
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
    Dim nMonthsL        As Long      'Strictly parsed nMonths
    Dim KeepEOM         As Boolean   'Strictly parsed Opt_KeepEOM
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

    'Reject a fractional, out-of-range or non-numeric nMonths
        If Not TryResolveLong(nMonths, nMonthsL, FailErr) Then GoTo Fail

    'Accept only an omitted, blank or native Boolean Opt_KeepEOM
        If Not TryResolveBool(Opt_KeepEOM, False, KeepEOM, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' COMPUTE SHIFT
'------------------------------------------------------------------------------
    'Delegate to the single month shifter; failure here is always range-related
        If Not TryAddMonths(ParsedDate, nMonthsL, KeepEOM, ResultDate) Then
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
    ByVal nYears As Variant, _
    Optional ByVal Opt_KeepEOM As Variant = False) _
    As Variant
'
'==============================================================================
' KPR_Dates_AddYears
'------------------------------------------------------------------------------
' PURPOSE
'   Adds nYears to DateIn by converting years to months (12 * nYears) and
'   applying the same EOM semantics as KPR_Dates_AddMonths.
'
' SIGNATURE
'   KPR_Dates_AddYears(DateIn, nYears, [Opt_KeepEOM]) -> Variant
'
' INPUTS
'   DateIn
'     Date candidate handled by the facade TryResolveDate policy.
'
'   nYears
'     Number of calendar years to add. Negative values are allowed.
'
'   Opt_KeepEOM (optional, default FALSE)
'     TRUE  => EOM in, EOM out
'     FALSE => clip day-of-month when the target month is shorter
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar, or unexpected runtime error
'   #NUM!    DateIn is outside the supported window
'   #NUM!    month delta or result falls outside the supported range
'
'   nYears is a Variant parsed strictly before conversion. Fractions and
'   Boolean values return #VALUE!; values outside the Long range return #NUM!.
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.TryAddMonths
'   - KPR_Core_Err.ErrValue, KPR_Core_Err.ErrNum
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
    Dim nYearsL         As Long      'Strictly parsed nYears
    Dim KeepEOM         As Boolean   'Strictly parsed Opt_KeepEOM
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
        If Not TryResolveDate(DateIn, ParsedDate, FailErr) Then GoTo Fail

    'Reject a fractional, out-of-range or non-numeric nYears
        If Not TryResolveLong(nYears, nYearsL, FailErr) Then GoTo Fail

    'Accept only an omitted, blank or native Boolean Opt_KeepEOM
        If Not TryResolveBool(Opt_KeepEOM, False, KeepEOM, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' CONVERT YEARS TO MONTHS
'------------------------------------------------------------------------------
    'Convert in Double to avoid intermediate overflow
        MonthsD = 12# * CDbl(nYearsL)

    'Gate to Long range before coercion
        If (MonthsD < -2147483648#) Or (MonthsD > 2147483647#) Then
            FailErr = ErrNum()
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' COMPUTE SHIFT
'------------------------------------------------------------------------------
    'Delegate to the single month shifter; failure here is always range-related
        If Not TryAddMonths(ParsedDate, CLng(MonthsD), KeepEOM, ResultDate) Then
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
    ByVal YearIn As Variant, _
    ByVal MonthIn As Variant, _
    ByVal WdIndex As Variant, _
    ByVal N As Variant, _
    Optional ByVal Opt_WeekBaseMonday As Variant = True) _
    As Variant
'
'==============================================================================
' KPR_Dates_NthWeekdayOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the N-th occurrence of a requested weekday within a given
'   Year / Month.
'
' SIGNATURE
'   KPR_Dates_NthWeekdayOfMonth(YearIn, MonthIn, WdIndex, N,
'                               [Opt_WeekBaseMonday]) -> Variant
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
'   N
'     Occurrence number in 1..5.
'
'   Opt_WeekBaseMonday (optional, default TRUE)
'     TRUE  => vbMonday base
'     FALSE => vbSunday base
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  YearIn, MonthIn, WdIndex or N outside their accepted domains,
'            or unexpected runtime error
'   #NUM!    arguments are valid but the occurrence does not exist in that
'            month (typically a requested 5th weekday), or the located date
'            falls outside KPR_MIN_DATE .. KPR_MAX_DATE
'
'   YearIn, MonthIn, WdIndex and N are Variants parsed strictly before their
'   function-specific domain checks. Fractions and Boolean values return
'   #VALUE!; values outside the Long range return #NUM!.
'
' DEPENDENCIES
'   - KPR_Core_Dates.DaysInMonth
'   - KPR_Core_Err.ErrValue, KPR_Core_Err.ErrNum
'   - VBA.DateSerial, VBA.Weekday
'
' NOTES
'   - The #VALUE! / #NUM! split is deliberate: "weekday 9" is a caller mistake,
'     "no 5th Friday in February" is a legitimate question with no answer.
'   - WdIndex must be passed consistently with Opt_WeekBaseMonday.
'   - Capping N at 5 is sufficient: no calendar month holds six occurrences of
'     the same weekday.
'   - The occurrence is located on serials and tested against the month length
'     BEFORE any date is constructed. Built as a Date first, a 5th occurrence
'     in December 9999 would step into year 10000 and raise error 5, which the
'     handler would report as #VALUE! rather than the #NUM! this case earns.
'   - This function takes a year and a month rather than a date, so it never
'     passes through TryResolveDate and never meets the shared window gate. It
'     therefore gates its own result. A year-granularity check is not enough:
'     the window floor is 1 March 1900, so January and February 1900 satisfy
'     the year test and would otherwise return an out-of-window date.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D0              As Date         'First calendar day of the target month
    Dim D0Serial        As Double       'Serial of D0, anchor for the occurrence walk
    Dim MonthLen        As Long         'Length in days of the target month
    Dim Off             As Long         'Days from D0 to the first occurrence of WdIndexL (0..6)
    Dim OutSerial       As Double       'Serial of the requested occurrence, before gating
    Dim WkBase          As VbDayOfWeek  'Weekday() base selector (vbMonday or vbSunday)
    Dim YearL           As Long      'Strictly parsed YearIn
    Dim MonthL          As Long      'Strictly parsed MonthIn
    Dim WdIndexL        As Long      'Strictly parsed WdIndex
    Dim NL              As Long      'Strictly parsed N
    Dim WkMonday        As Boolean   'Strictly parsed Opt_WeekBaseMonday
    Dim FailErr         As Variant      'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

    'Reject a fractional, out-of-range or non-numeric YearIn
        If Not TryResolveLong(YearIn, YearL, FailErr) Then GoTo Fail

    'Reject a fractional, out-of-range or non-numeric MonthIn
        If Not TryResolveLong(MonthIn, MonthL, FailErr) Then GoTo Fail

    'Reject a fractional, out-of-range or non-numeric WdIndex
        If Not TryResolveLong(WdIndex, WdIndexL, FailErr) Then GoTo Fail

    'Reject a fractional, out-of-range or non-numeric N
        If Not TryResolveLong(N, NL, FailErr) Then GoTo Fail

    'Accept only an omitted, blank or native Boolean Opt_WeekBaseMonday
        If Not TryResolveBool(Opt_WeekBaseMonday, True, WkMonday, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Require a year inside the module supported window
        If (YearL < Year(KPR_MIN_DATE)) Or (YearL > Year(KPR_MAX_DATE)) Then
            FailErr = ErrForCondition(KPR_COND_DOMAIN_YEAR)
            GoTo Fail
        End If

    'Require a valid calendar month
        If (MonthL < 1) Or (MonthL > 12) Then
            FailErr = ErrForCondition(KPR_COND_DOMAIN_MONTH)
            GoTo Fail
        End If

    'Require a weekday index in 1..7 under the selected base
        If (WdIndexL < 1) Or (WdIndexL > 7) Then
            FailErr = ErrForCondition(KPR_COND_DOMAIN_WEEKDAY)
            GoTo Fail
        End If

    'Require an occurrence number in 1..5
        If (NL < 1) Or (NL > 5) Then
            FailErr = ErrForCondition(KPR_COND_DOMAIN_OCCURRENCE)
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' RESOLVE WEEKDAY BASE
'------------------------------------------------------------------------------
    'Resolve Weekday() base once
        If WkMonday Then
            WkBase = vbMonday
        Else
            WkBase = vbSunday
        End If

'------------------------------------------------------------------------------
' COMPUTE REQUESTED OCCURRENCE
'------------------------------------------------------------------------------
    'Anchor at the first day of the requested month
        D0 = DateSerial(YearL, MonthL, 1)
        D0Serial = CDbl(D0)

    'Length of the target month, read without constructing its month-end
        MonthLen = DaysInMonth(YearL, MonthL)

    'Forward offset (0..6) from the anchor to the first requested weekday
        Off = (WdIndexL - Weekday(D0, WkBase) + 7) Mod 7

    'Walk to the requested occurrence on serials, not on dates
        OutSerial = D0Serial + CDbl(Off) + (7# * CDbl(NL - 1))

    'Reject an occurrence that would spill past the end of the month
        If OutSerial > (D0Serial + CDbl(MonthLen) - 1#) Then
            FailErr = ErrNum()
            GoTo Fail
        End If

    'Reject a located date outside the supported window
        If (OutSerial < CDbl(KPR_MIN_DATE)) Or (OutSerial > CDbl(KPR_MAX_DATE)) Then
            FailErr = ErrNum()
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the requested occurrence
        KPR_Dates_NthWeekdayOfMonth = CDate(OutSerial)
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
    ByVal YearIn As Variant, _
    ByVal MonthIn As Variant, _
    ByVal WdIndex As Variant, _
    Optional ByVal Opt_WeekBaseMonday As Variant = True) _
    As Variant
'
'==============================================================================
' KPR_Dates_LastWeekdayOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last occurrence of a requested weekday within a given
'   Year / Month.
'
' SIGNATURE
'   KPR_Dates_LastWeekdayOfMonth(YearIn, MonthIn, WdIndex,
'                                [Opt_WeekBaseMonday]) -> Variant
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
'   Opt_WeekBaseMonday (optional, default TRUE)
'     TRUE  => vbMonday base
'     FALSE => vbSunday base
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  YearIn, MonthIn or WdIndex outside their accepted domains, or
'            unexpected runtime error
'   #NUM!    the located date falls outside KPR_MIN_DATE .. KPR_MAX_DATE
'
'   YearIn, MonthIn and WdIndex are Variants parsed strictly before their
'   function-specific domain checks. Fractions and Boolean values return
'   #VALUE!; values outside the Long range return #NUM!.
'
' DEPENDENCIES
'   - KPR_Core_Dates.DaysInMonth
'   - KPR_Core_Err.ErrValue, KPR_Core_Err.ErrNum
'   - VBA.DateSerial, VBA.Weekday
'
' NOTES
'   - A last occurrence always exists inside the month, so the only #NUM! path
'     is the window gate.
'   - WdIndex must be passed consistently with Opt_WeekBaseMonday.
'   - The month-end anchor is built from DaysInMonth rather than from the day-0
'     idiom DateSerial(y, m + 1, 0). That idiom raises error 5 for December
'     9999, which is inside the supported window, and the handler would report
'     it as #VALUE! rather than returning the correct date.
'   - This function takes a year and a month rather than a date, so it never
'     passes through TryResolveDate and never meets the shared window gate. It
'     therefore gates its own result. A year-granularity check is not enough:
'     the window floor is 1 March 1900, so January and February 1900 satisfy
'     the year test and would otherwise return an out-of-window date.
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
    Dim OutSerial       As Double       'Serial of the located date, before gating
    Dim WkBase          As VbDayOfWeek  'Weekday() base selector (vbMonday or vbSunday)
    Dim YearL           As Long      'Strictly parsed YearIn
    Dim MonthL          As Long      'Strictly parsed MonthIn
    Dim WdIndexL        As Long      'Strictly parsed WdIndex
    Dim WkMonday        As Boolean   'Strictly parsed Opt_WeekBaseMonday
    Dim FailErr         As Variant      'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

    'Reject a fractional, out-of-range or non-numeric YearIn
        If Not TryResolveLong(YearIn, YearL, FailErr) Then GoTo Fail

    'Reject a fractional, out-of-range or non-numeric MonthIn
        If Not TryResolveLong(MonthIn, MonthL, FailErr) Then GoTo Fail

    'Reject a fractional, out-of-range or non-numeric WdIndex
        If Not TryResolveLong(WdIndex, WdIndexL, FailErr) Then GoTo Fail

    'Accept only an omitted, blank or native Boolean Opt_WeekBaseMonday
        If Not TryResolveBool(Opt_WeekBaseMonday, True, WkMonday, FailErr) Then GoTo Fail

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Require a year inside the module supported window
        If (YearL < Year(KPR_MIN_DATE)) Or (YearL > Year(KPR_MAX_DATE)) Then
            FailErr = ErrForCondition(KPR_COND_DOMAIN_YEAR)
            GoTo Fail
        End If

    'Require a valid calendar month
        If (MonthL < 1) Or (MonthL > 12) Then
            FailErr = ErrForCondition(KPR_COND_DOMAIN_MONTH)
            GoTo Fail
        End If

    'Require a weekday index in 1..7 under the selected base
        If (WdIndexL < 1) Or (WdIndexL > 7) Then
            FailErr = ErrForCondition(KPR_COND_DOMAIN_WEEKDAY)
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' RESOLVE WEEKDAY BASE
'------------------------------------------------------------------------------
    'Resolve Weekday() base once
        If WkMonday Then
            WkBase = vbMonday
        Else
            WkBase = vbSunday
        End If

'------------------------------------------------------------------------------
' COMPUTE LAST OCCURRENCE
'------------------------------------------------------------------------------
    'Anchor at month-end, built inside the month rather than by rolling forward
        EOM = DateSerial(YearL, MonthL, DaysInMonth(YearL, MonthL))

    'Backward offset (0..6) from the anchor to the requested weekday
        Diff = (Weekday(EOM, WkBase) - WdIndexL + 7) Mod 7

    'Step back on serials; the result always stays inside the month
        OutSerial = CDbl(EOM) - CDbl(Diff)

    'Reject a located date outside the supported window
        If (OutSerial < CDbl(KPR_MIN_DATE)) Or (OutSerial > CDbl(KPR_MAX_DATE)) Then
            FailErr = ErrNum()
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the last requested weekday in the target month
        KPR_Dates_LastWeekdayOfMonth = CDate(OutSerial)
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
' SIGNATURE
'   KPR_Dates_PillarFromDates(StartDate, EndDate) -> Variant
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
'   #VALUE!  either date not parseable / not scalar / outside the supported
'            window / unexpected runtime error
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.Pillar_Format_Nearest
'   - KPR_Core_Err.ErrValue
'
' NOTES
'   - The label is a String by design; it is a category, not a quantity, and is
'     meant to be grouped or matched rather than computed on.
'   - EndDate before StartDate yields a "-" prefixed token.
'   - The week family is capped at 3W, so no label of 4W or longer is emitted
'     and the 3W to 1M boundary falls at 25 days. Labels are nearest tenor
'     names, not members of a quoted pillar set: any integer month can appear,
'     including months a desk would not quote.
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
        If Not TryResolveDate(StartDate, ParsedStart, FailErr) Then GoTo Fail

    'Parse the end date under module policy
        If Not TryResolveDate(EndDate, ParsedEnd, FailErr) Then GoTo Fail

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
' SIGNATURE
'   KPR_Dates_DatesFromPillar(StartDate, Pillar) -> Variant
'
' INPUTS
'   StartDate
'     Date candidate handled by the facade TryResolveDate policy.
'
'   Pillar
'     Tenor token, unwrapped by the same shape policy as StartDate. Accepted
'     grammar:
'       - optional leading sign: "+" or "-", applied to the WHOLE token
'       - one or more [integer][unit] components, no spaces inside
'       - units: Y, M, W, D (case insensitive), each at most once
'       - whole-token aliases: "ON" / "O/N" => 1D, "TN" / "T/N" => 2D,
'         never signed
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
'   - KPR_Core_Array.TryUnwrapScalar
'   - KPR_Core_Dates.TryPillar_Parse
'   - KPR_Core_Dates.TryAddMonths
'   - KPR_Core_Err.ErrValue, KPR_Core_Err.ErrNum
'
' NOTES
'   - Y and M components are aggregated into one calendar-month shift, applied
'     first with CLIP semantics (never preserve-EOM). W and D are then applied
'     as an exact calendar-day delta. Order matters: "1M1D" from 31-Jan-2026 is
'     28-Feb + 1D = 01-Mar-2026.
'   - A repeated unit is rejected, not summed: "1Y2Y3M" fails rather than
'     resolving to 3Y3M. A duplicated unit is a typo, and returning a plausible
'     date for it would be worse than refusing.
'   - The pillar argument goes through TryUnwrapScalar, so a single-cell Range
'     or a 1x1 wrapper is accepted exactly as it is for StartDate. Rejecting
'     wrappers here while accepting them there would give the two arguments
'     different shape policies in one signature.
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
    Dim PillarScalar    As Variant   'Pillar payload after shape unwrapping
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
        If Not TryResolveDate(StartDate, ParsedStart, FailErr) Then GoTo Fail

    'Reduce the pillar argument under the same shape policy as StartDate
        If Not TryUnwrapScalar(Pillar, PillarScalar) Then GoTo Fail

    'Parse the pillar token into signed month / day deltas
        If Not TryPillar_Parse(PillarScalar, TotalMonths, TotalDays) Then GoTo Fail

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
                If Not TryAddMonths(WorkDate, CLng(TotalMonths), False, WorkDate) Then
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
    ByRef ParsedDate As Date, _
    ByRef ErrOut As Variant) _
    As Boolean
'
'==============================================================================
'                                TryResolveDate
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves one worksheet date argument into a normalized, in-window VBA Date,
'   or assigns the exact worksheet error the contract requires.
'
' SIGNATURE
'   TryResolveDate(DateIn, ParsedDate, ErrOut) -> Boolean
'
' OUTPUTS
'   ParsedDate (ByRef)   assigned ONLY on success
'   ErrOut     (ByRef)   assigned ONLY on failure, to the value to return
'
' BEHAVIOR
'   1. KPR_Core_Array.TryUnwrapScalar reduces wrappers to one scalar. Failure is
'      SHAPE_UNSUPPORTED.
'   2. An incoming native Excel error is returned VERBATIM. This is the
'      public/element propagation boundary: the error never reaches the parser
'      and is never collapsed into an ordinary parse failure.
'   3. KPR_Core_Parse.TryParseDateScalar classifies and window-gates the value.
'   4. The condition is mapped through KPR_Core_Err.ErrForCondition, so a
'      malformed value yields #VALUE! and an out-of-window value yields #NUM!.
'
' NOTES
'   - Propagation is deliberately not routed through ErrForCondition. The value
'     to return is the incoming error itself, not an error chosen for a
'     condition, and ErrForCondition raises rather than mapping propagation.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScalarValue     As Variant          'Payload after shape unwrapping
    Dim Candidate       As Date             'Parsed candidate
    Dim Condition       As KPR_Condition    'Classified failure condition

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default return is failure unless every stage succeeds
        TryResolveDate = False

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'Reduce a single-cell Range or 1x1 wrapper to its scalar payload
        If Not TryUnwrapScalar(DateIn, ScalarValue) Then
            ErrOut = ErrForCondition(KPR_COND_SHAPE_UNSUPPORTED)
            Exit Function
        End If

'------------------------------------------------------------------------------
' PROPAGATE INCOMING ERRORS
'------------------------------------------------------------------------------
    'An incoming Excel error is the caller's answer, returned unchanged
        If VarType(ScalarValue) = vbError Then
            ErrOut = ScalarValue
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'Classify strictly and apply the supported window
        If Not TryParseDateScalar(ScalarValue, Candidate, Condition) Then
            ErrOut = ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        ParsedDate = Candidate

    'Contract: TRUE only when ParsedDate was assigned
        TryResolveDate = True

End Function

Private Function TryResolveLong( _
    ByVal ValueIn As Variant, _
    ByRef ParsedLong As Long, _
    ByRef ErrOut As Variant) _
    As Boolean
'
'==============================================================================
'                                TryResolveLong
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves one worksheet integer argument into a Long, or assigns the exact
'   worksheet error the contract requires.
'
' BEHAVIOR
'   Same four stages as TryResolveDate, with KPR_Core_Parse.TryParseLongScalar
'   doing the classification. A fraction yields #VALUE! under INTEGER_FRACTION
'   and a value outside the Long range yields #NUM! under INTEGER_RANGE.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScalarValue     As Variant          'Payload after shape unwrapping
    Dim Candidate       As Long             'Parsed candidate
    Dim Condition       As KPR_Condition    'Classified failure condition

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default return is failure unless every stage succeeds
        TryResolveLong = False

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'Reduce a single-cell Range or 1x1 wrapper to its scalar payload
        If Not TryUnwrapScalar(ValueIn, ScalarValue) Then
            ErrOut = ErrForCondition(KPR_COND_SHAPE_UNSUPPORTED)
            Exit Function
        End If

'------------------------------------------------------------------------------
' PROPAGATE INCOMING ERRORS
'------------------------------------------------------------------------------
    'An incoming Excel error is the caller's answer, returned unchanged
        If VarType(ScalarValue) = vbError Then
            ErrOut = ScalarValue
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'Classify strictly: no truncation, no rounding, no Boolean coercion
        If Not TryParseLongScalar(ScalarValue, Candidate, Condition) Then
            ErrOut = ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        ParsedLong = Candidate

    'Contract: TRUE only when ParsedLong was assigned
        TryResolveLong = True

End Function

Private Function TryResolveBool( _
    ByVal ControlIn As Variant, _
    ByVal DefaultValue As Boolean, _
    ByRef ParsedBool As Boolean, _
    ByRef ErrOut As Variant) _
    As Boolean
'
'==============================================================================
'                                TryResolveBool
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves one optional Boolean control, or assigns the worksheet error the
'   contract requires.
'
' BEHAVIOR
'   - Each public Optional Variant declaration supplies its documented default,
'     so an omitted argument reaches this helper as a native Boolean.
'   - A blank cell arrives as Empty and selects the same default.
'   - A control larger than 1x1 is CONTROL_NOT_SCALAR in the contract; this
'     revision reports the shape rejection it inherits from TryUnwrapScalar,
'     which yields the same #VALUE!. Issue #16 separates the two when arrays
'     become traversable.
'   - An incoming error in a control propagates verbatim as a call-level result.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScalarValue     As Variant          'Payload after shape unwrapping
    Dim Candidate       As Boolean          'Parsed candidate
    Dim Condition       As KPR_Condition    'Classified failure condition

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default return is failure unless every stage succeeds
        TryResolveBool = False

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'A control must be scalar or a 1x1 wrapper; it never vectorizes
        If Not TryUnwrapScalar(ControlIn, ScalarValue) Then
            ErrOut = ErrForCondition(KPR_COND_SHAPE_UNSUPPORTED)
            Exit Function
        End If

'------------------------------------------------------------------------------
' PROPAGATE INCOMING ERRORS
'------------------------------------------------------------------------------
    'An incoming Excel error is the caller's answer, returned unchanged
        If VarType(ScalarValue) = vbError Then
            ErrOut = ScalarValue
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'Accept a native Boolean only; Empty selects the documented default
        If Not TryParseBoolControl(ScalarValue, DefaultValue, Candidate, Condition) Then
            ErrOut = ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the control is validated
        ParsedBool = Candidate

    'Contract: TRUE only when ParsedBool was assigned
        TryResolveBool = True

End Function
