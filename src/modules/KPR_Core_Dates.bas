Attribute VB_Name = "KPR_Core_Dates"
'==============================================================================
' MODULE: KPR_Core_Dates
'------------------------------------------------------------------------------
' PURPOSE
'   Internal boundary for calendar computation: Gregorian arithmetic, month
'   boundary logic, EOM-aware month shifting, and the pillar parse/format core.
'
'   This module owns the supported date window, because the window is a
'   calendar fact rather than a parsing rule.
'
' SCOPE
'   - Supported window constants and the window predicate
'   - IsLeapYear
'   - DaysInMonth
'   - EndOfMonth
'   - TryAddMonths
'   - TryPillar_Parse
'   - Pillar_Format_Nearest
'
' VISIBILITY
'   Option Private Module. Members are declared Public so other modules in the
'   project can call them; the module-level Private setting is what keeps them
'   out of the Excel function list and off the project's external surface.
'   Nothing here is supported API.
'
' ALLOWED DEPENDENCIES
'   None. This module computes on already-validated inputs and never parses,
'   never inspects shape, and never constructs a worksheet error. VBA date
'   intrinsics are used directly.
'
' SUPPORTED WINDOW
'       KPR_MIN_DATE = 1900-03-01
'       KPR_MAX_DATE = 9999-12-31
'
'   The lower bound is serial 61 by design: it excludes the Excel 1900-system
'   region where worksheet serials and VBA dates disagree because of the
'   fictitious 29-Feb-1900. At and above serial 61 the two systems agree.
'
'   The baseline applied this gate inside its combined parser. The gate is now
'   applied by the facade immediately after parsing, using the predicate below.
'   It remains a single gate on every accepted path; only its location moved,
'   so that KPR_Core_Parse need not depend on this module.
'
'   The derived index constants below restate the window in absolute months.
'   They exist to keep per-call arithmetic off the hot paths and MUST be
'   updated together with the two date constants.
'
' MONTH LENGTH
'   DaysInMonth is the single source of month length for this module, and it
'   is the only member that consults the leap rule. EndOfMonth is built from
'   it, and TryAddMonths reads it directly.
'
'   The day-0 idiom DateSerial(y, m + 1, 0) is deliberately NOT used anywhere
'   in this module. It raises error 5 at both ends of the VBA Date range:
'   DateSerial(9999, 13, 0) and DateSerial(100, 1, 0) both fail, verified in
'   isolation on the host. Deriving month length arithmetically removes the
'   boundary crossing rather than guarding it.
'
' SCOPE BOUNDARY (THIS REVISION)
'   Pillar rounding is NEAREST only, as migrated. FLOOR and CEILING arrive with
'   Opt_Rounding in issue #14, which owns the rounding modes and the accepted
'   token grammar together.
'
'   Duplicate pillar units and signed aliases are already rejected here, matching
'   the contract grammar. Both rejections signal a bare parse failure: the
'   PILLAR_DUPLICATE_UNIT and PILLAR_ALIAS_SIGNED identifiers the contract
'   registers for them do not exist in KPR_Condition yet, so a caller cannot
'   tell the two apart or distinguish either from a malformed token. Issue #14
'   closes that gap when it takes ownership of the grammar; this module cannot
'   do it alone, because the dependency matrix does not let it reach the
'   condition vocabulary in KPR_Core_Err.
'
' NOTES
'   - Date math follows Excel / VBA DateSerial, DateAdd, DateDiff and Weekday
'     semantics. Gregorian leap-year logic is used throughout.
'   - Day distances between two known date-only values are taken by serial
'     subtraction rather than DateDiff. The results are identical because both
'     operands are integral, and the subtraction is materially cheaper on the
'     per-cell paths.
'   - The facade AddYears delegates to TryAddMonths so year-roll edge cases
'     (29-Feb, short months, EOM) cannot drift between the two surfaces.
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
    Option Explicit         'Force explicit variable declarations
    Option Private Module   'Internal module: invisible outside this VBA project

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Supported date window (single source of truth for every parse / result gate)
        Public Const KPR_MIN_DATE   As Date = #3/1/1900#     'Serial 61; excludes the Excel 1900 leap-bug region
        Public Const KPR_MAX_DATE   As Date = #12/31/9999#   'Serial 2958465; upper bound of the Excel date system

    'The same window as absolute 0-based month indices, so the hot paths do not
    'recompute Year / Month of a constant on every call. Keep in step with the
    'two date constants above.
        Private Const KPR_MIN_YEAR      As Long = 1900       'Year(KPR_MIN_DATE)
        Private Const KPR_MAX_YEAR      As Long = 9999       'Year(KPR_MAX_DATE)
        Private Const KPR_MAX_MONTHIDX  As Long = 119999     '9999 * 12 + 12 - 1

'
'------------------------------------------------------------------------------
'
'                              SUPPORTED WINDOW
'
'------------------------------------------------------------------------------
'

Public Function IsDateInWindow( _
    ByVal DateIn As Date) _
    As Boolean
'
'==============================================================================
'                                IsDateInWindow
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a parsed date lies inside the supported window.
'
' SIGNATURE
'   IsDateInWindow(DateIn) -> Boolean
'
' INPUTS
'   DateIn
'     Any VBA Date value.
'
' RETURNS
'   Boolean
'     TRUE  => KPR_MIN_DATE <= DateIn <= KPR_MAX_DATE
'     FALSE => outside the supported window
'
' ERROR POLICY
'   None. The caller decides which worksheet error an out-of-window value maps
'   to.
'
' NOTES
'   - The migrated surface maps a window failure to #VALUE!, which is baseline
'     behaviour. The contract classifies it as DATE_WINDOW and requires #NUM!;
'     that change is owned by issues #12 and #13.
'   - A value carrying a time component is compared as a serial, so 9999-12-31
'     with a non-zero time reports FALSE. Callers holding a raw parse result
'     should floor it before gating if that is not intended.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'One window test for every accepted path
        IsDateInWindow = ((DateIn >= KPR_MIN_DATE) And (DateIn <= KPR_MAX_DATE))

End Function

'
'------------------------------------------------------------------------------
'
'                            CALENDAR COMPUTATION CORE
'
'------------------------------------------------------------------------------
'

Public Function IsLeapYear( _
    ByVal YearIn As Long) _
    As Boolean
'
'==============================================================================
'                                  IsLeapYear
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the Gregorian leap-year rule to a calendar year.
'
' WHY THIS EXISTS
'   The rule is short enough to inline and subtle enough to get wrong. Stating
'   the century exceptions once means DaysInMonth is the only place February
'   is resolved, and every other month length derives from that.
'
' SIGNATURE
'   IsLeapYear(YearIn) -> Boolean
'
' INPUTS
'   YearIn
'     Calendar year, not a date serial. Callers holding a Date must pass
'     Year(DateIn).
'
' RETURNS
'   Boolean
'     TRUE => YearIn is a leap year under the proleptic Gregorian rule.
'
' ERROR POLICY
'   None. Total for every Long; the caller owns range validation.
'
' NOTES
'   - Rule: divisible by 4 => leap, unless divisible by 100, unless also
'     divisible by 400. So 1900 is not a leap year and 2000 is.
'   - The rule is applied proleptically. VBA Date arithmetic is Gregorian
'     over its whole range, so no Julian calendar or 1582 changeover is
'     modelled here and 1500 reports TRUE. That is consistent with the rest
'     of the layer, not a historical claim.
'   - The supported window is not enforced. Any Long is evaluated, including
'     zero and negatives; VBA Mod takes the sign of the dividend, so -4
'     reports TRUE. Values outside the window are the caller's to reject.
'
' UPDATED
'   2026-08-31
'==============================================================================
'
'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Gregorian rule
        IsLeapYear = ((YearIn Mod 4 = 0) And ((YearIn Mod 100 <> 0) Or (YearIn Mod 400 = 0)))
End Function


Public Function DaysInMonth( _
    ByVal YearIn As Long, _
    ByVal MonthIn As Long) _
    As Long
'
'==============================================================================
'                                  DaysInMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the length in days of a given calendar month.
'
' WHY THIS EXISTS
'   Month length is needed by EndOfMonth, by the EOM anchor test, and by the
'   day clip in TryAddMonths. Deriving it arithmetically rather than through
'   DateSerial keeps every one of those paths inside the Date range: the day-0
'   idiom raises error 5 in December 9999 and in January of the floor year.
'
' SIGNATURE
'   DaysInMonth(YearIn, MonthIn) -> Long
'
' INPUTS
'   YearIn
'     Calendar year, not a date serial. Read only for February.
'
'   MonthIn
'     Calendar month, 1 to 12.
'
' RETURNS
'   Long
'     28, 29, 30 or 31 for a valid month; 0 for any MonthIn outside 1 to 12.
'
' ERROR POLICY
'   None. Total for every pair of Longs, including an invalid month.
'
' DEPENDENCIES
'   - IsLeapYear
'
' NOTES
'   - A zero return signals an invalid month rather than raising, because this
'     is an internal member reached only with an already-resolved month. A
'     caller that can produce an out-of-range month must test the result; one
'     that cannot may use it directly.
'   - No date is constructed, so the result is valid for any year the caller
'     can represent, including years outside the supported window.
'
' UPDATED
'   2026-08-31
'==============================================================================
'
'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Fixed lengths first; February is the only month needing the leap rule
        Select Case MonthIn
            Case 1, 3, 5, 7, 8, 10, 12
                DaysInMonth = 31
            Case 4, 6, 9, 11
                DaysInMonth = 30
            Case 2
                If IsLeapYear(YearIn) Then
                    DaysInMonth = 29
                Else
                    DaysInMonth = 28
                End If
            Case Else
                'Invalid month; reported rather than raised
                    DaysInMonth = 0
        End Select
End Function


Public Function EndOfMonth( _
    ByVal DateIn As Date) _
    As Date
'
'==============================================================================
'                                  EndOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last calendar day of the month containing DateIn.
'
' WHY THIS EXISTS
'   Month-end is the anchor for the whole month-arithmetic family. Isolating
'   it here means no caller re-derives it.
'
' SIGNATURE
'   EndOfMonth(DateIn) -> Date
'
' INPUTS
'   DateIn
'     Already-parsed VBA Date. A time component is tolerated and discarded.
'
' RETURNS
'   Date
'     Midnight on the last calendar day of the containing month.
'
' ERROR POLICY
'   None. Total over the whole VBA Date range, both boundary months included.
'
' DEPENDENCIES
'   - DaysInMonth
'
' NOTES
'   - The month length comes from DaysInMonth, so no date outside the month
'     is ever constructed. The day-0 idiom would be shorter but raises error 5
'     in December 9999 and in January of the floor year; see MONTH LENGTH in
'     the module header.
'   - Year and Month ignore any time fraction, and DateSerial always returns
'     midnight, so the result is date-only regardless of what arrives. No
'     caller needs to strip the time first.
'   - Totality is asserted over the whole VBA Date range, which is wider than
'     the supported window. Pre-1900 dates are negative serials and behave
'     identically.
'
' UPDATED
'   2026-08-31
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim YearPart    As Long     'Calendar year of DateIn
    Dim MonthPart   As Long     'Calendar month of DateIn, 1 to 12
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Decompose once; both parts feed the month-length lookup
        YearPart = Year(DateIn)
        MonthPart = Month(DateIn)
'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Last day of the containing month, built without leaving it
        EndOfMonth = DateSerial(YearPart, MonthPart, DaysInMonth(YearPart, MonthPart))
End Function


Public Function TryAddMonths( _
    ByVal DateIn As Date, _
    ByVal nMonths As Long, _
    ByVal KeepEOM As Boolean, _
    ByRef DateOut As Date) _
    As Boolean
'
'==============================================================================
'                                 TryAddMonths
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
'   TryAddMonths(DateIn, nMonths, KeepEOM, DateOut) -> Boolean
'
' INPUTS
'   DateIn
'     Already-parsed VBA Date. A time component is tolerated and discarded.
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
'   - Clears Err before returning FALSE, so no stale error state survives.
'
' DEPENDENCIES
'   - DaysInMonth
'
' NOTES
'   - DateOut is safe to pass as the same variable as an input date at the call
'     site, because it is written only once, at the end.
'   - Both month lengths come from DaysInMonth, so neither the EOM anchor test
'     nor the day clip constructs a date outside the month it is measuring.
'   - The index gate is deliberately coarse, admitting whole years at both
'     ends. The terminal gate is what enforces the exact window, so if either
'     bound ever moves off a year boundary the pair still holds.
'   - The index is computed in Double so that an extreme nMonths cannot
'     overflow before the gate can reject it.
'
' UPDATED
'   2026-08-31
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
    Dim TargetD         As Long      'Resolved day-of-month in the target month

    Dim Candidate       As Date      'Shifted date before the terminal range gate

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime errors and convert them to a FALSE return
        On Error GoTo Fail
    'Default return is failure unless we explicitly succeed
        TryAddMonths = False

'------------------------------------------------------------------------------
' CACHE INPUT COMPONENTS
'------------------------------------------------------------------------------
    'Split the input date once
        Y = Year(DateIn)
        M = Month(DateIn)
        DayIn = Day(DateIn)
    'Detect the EOM anchor before shifting
        IsEOM = (DayIn = DaysInMonth(Y, M))

'------------------------------------------------------------------------------
' RESOLVE TARGET MONTH
'------------------------------------------------------------------------------
    'Absolute 0-based month index, shifted, computed in Double
        MonthIndex = (CDbl(Y) * 12# + CDbl(M) - 1#) + CDbl(nMonths)
    'Gate the index before any coercion so DateSerial never sees a wild value
        If (MonthIndex < (CDbl(KPR_MIN_YEAR) * 12#)) Or _
           (MonthIndex > (CDbl(KPR_MAX_YEAR) * 12# + 11#)) Then GoTo Fail
    'Split the gated index back into year and month
        TargetY = CLng(Int(MonthIndex / 12#))
        TargetM = CLng(MonthIndex - (CDbl(TargetY) * 12#)) + 1

'------------------------------------------------------------------------------
' RESOLVE DAY OF MONTH
'------------------------------------------------------------------------------
    'Target month length, read without constructing a date
        TargetDim = DaysInMonth(TargetY, TargetM)

    'Preserve EOM when asked and the input was EOM; otherwise clip
        If KeepEOM And IsEOM Then
            TargetD = TargetDim
        Else
            TargetD = DayIn
            If TargetD > TargetDim Then TargetD = TargetDim
        End If

'------------------------------------------------------------------------------
' TERMINAL RANGE GATE
'------------------------------------------------------------------------------
    'Build the candidate and gate it against the supported window
        Candidate = DateSerial(TargetY, TargetM, TargetD)
        If (Candidate < KPR_MIN_DATE) Or (Candidate > KPR_MAX_DATE) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        DateOut = Candidate
    'Contract: TRUE only when DateOut was assigned
        TryAddMonths = True
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving DateOut untouched
        TryAddMonths = False
    'Do not leave stale error state for the caller to observe
        Err.Clear

End Function

'
'------------------------------------------------------------------------------
'
'                                 PILLAR CORE
'
'------------------------------------------------------------------------------
'

Public Function TryPillar_Parse( _
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
'     FALSE => blank, malformed, unknown unit, repeated unit, signed alias,
'              or non-text payload
'
' BEHAVIOR
'   - Trims and upper-cases, consumes an optional leading sign, then checks the
'     whole-token aliases ON / O/N / TN / T/N.
'   - Otherwise scans repeated [digits][unit] pairs to the end of the string.
'   - Any character that is not a digit in a numeric position or a known unit
'     in a unit position fails the parse.
'   - Each unit may appear at most once, so "1M1M" fails rather than summing.
'
' ERROR POLICY
'   - Does not raise. All failure paths return FALSE.
'   - Clears Err before returning FALSE, so no stale error state survives.
'
' DEPENDENCIES
'   None.
'
' NOTES
'   - No characters are stripped from the body, so "1 M" and "1/M" fail rather
'     than being silently reinterpreted as "1M". The aliases are the only
'     tokens containing "/", and they are matched whole.
'   - Fractional tenors are not part of the grammar: "1.5M" fails. Half months
'     are not a market convention, and admitting them would force a rounding
'     policy into a parser.
'   - A leading sign on an alias is rejected. The aliases name fixed points at
'     the short end, not quantities, so "-ON" has no meaning to negate.
'   - A repeated unit is a typo, not a sum. Accepting "1M1M" as two months
'     would return a plausible number for input the caller did not intend.
'     This matches the contract; the change is owned by issue #15.
'   - Quantities are not bounded here. A token such as "999999999999M" parses
'     and the magnitude is left for the shift layer to reject, which it does
'     at its month-index gate.
'   - CDbl is applied to a digits-only substring, so no locale decimal
'     separator is involved. Replacing it with Val or CLng would change the
'     overflow behaviour.
'   - The digit scan compares character codes rather than one-character
'     strings. The unit test stays on strings because it runs once per
'     component rather than once per character.
'   - A signed result can be negative zero, because the sign multiplier is
'     applied to a zero aggregate. It compares equal to zero and behaves
'     identically in arithmetic; only a formatter would notice.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S               As String    'Upper-cased, trimmed pillar text
    Dim SBody           As String    'Pillar body after the optional leading sign
    Dim ChCode          As Long      'Character code at the current scan position
    Dim HasSign         As Boolean   'TRUE when a leading sign was consumed

    Dim ScanPos         As Long      'Current scan position in SBody
    Dim TokenStart      As Long      'Start position of the current numeric token
    Dim TokenCount      As Long      'Number of components parsed so far
    Dim BodyLen         As Long      'Length of SBody, read once

    Dim SignMul         As Double    'Global sign multiplier (+1 or -1)
    Dim QtyD            As Double    'Current component quantity

    Dim SeenY           As Boolean   'TRUE once a Y component has been consumed
    Dim SeenM           As Boolean   'TRUE once an M component has been consumed
    Dim SeenW           As Boolean   'TRUE once a W component has been consumed
    Dim SeenD           As Boolean   'TRUE once a D component has been consumed

    Dim YearsD          As Double    'Y component
    Dim MonthsD         As Double    'M component
    Dim WeeksD          As Double    'W component
    Dim DaysD           As Double    'D component, or the resolved alias

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
                HasSign = True
                SBody = Mid$(S, 2)
            Case "-"
                HasSign = True
                SignMul = -1#
                SBody = Mid$(S, 2)
            Case Else
                SBody = S
        End Select

    'Reject a sign with no body
        BodyLen = Len(SBody)
        If BodyLen = 0 Then GoTo Fail

'------------------------------------------------------------------------------
' WHOLE-TOKEN ALIASES
'------------------------------------------------------------------------------
    'Overnight and tom-next are matched whole, never as generic tokens
        Select Case SBody
            Case "ON", "O/N", "TN", "T/N"
                'An alias names a fixed point and cannot carry a sign
                    If HasSign Then GoTo Fail
                'Resolve the alias to its day count
                    If (SBody = "ON") Or (SBody = "O/N") Then
                        DaysD = 1#
                    Else
                        DaysD = 2#
                    End If
                    TokenCount = 1
        End Select

'------------------------------------------------------------------------------
' GENERIC INTEGER + UNIT SCANNER
'------------------------------------------------------------------------------
    'Only run the scanner when no alias matched
        If TokenCount = 0 Then

            'Start at the first character of the body
                ScanPos = 1
            'Consume [digits][unit] pairs until the body is exhausted
                Do While ScanPos <= BodyLen
                    'Mark the start of the numeric token
                        TokenStart = ScanPos
                    'Advance through contiguous digits, comparing codes
                        Do While ScanPos <= BodyLen
                            ChCode = AscW(Mid$(SBody, ScanPos, 1))
                            If (ChCode < 48) Or (ChCode > 57) Then Exit Do
                            ScanPos = ScanPos + 1
                        Loop
                    'Reject a component with no digits
                        If TokenStart = ScanPos Then GoTo Fail
                    'Coerce the numeric token once
                        QtyD = CDbl(Mid$(SBody, TokenStart, ScanPos - TokenStart))
                    'Reject a quantity with no trailing unit
                        If ScanPos > BodyLen Then GoTo Fail
                    'Record the component by unit, rejecting any repeat
                        Select Case Mid$(SBody, ScanPos, 1)
                            Case "Y"
                                If SeenY Then GoTo Fail
                                SeenY = True
                                YearsD = QtyD
                            Case "M"
                                If SeenM Then GoTo Fail
                                SeenM = True
                                MonthsD = QtyD
                            Case "W"
                                If SeenW Then GoTo Fail
                                SeenW = True
                                WeeksD = QtyD
                            Case "D"
                                If SeenD Then GoTo Fail
                                SeenD = True
                                DaysD = QtyD
                            Case Else
                                GoTo Fail
                        End Select
                    'Count the component and step past the unit
                        TokenCount = TokenCount + 1
                        ScanPos = ScanPos + 1
                Loop

        End If

    'Unreachable in practice: the scanner fails before exiting with no
    'components, and a blank body was rejected earlier. Kept as a guard so
    'the TRUE return can never be reached without a parsed component.
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
    'Do not leave stale error state for the caller to observe
        Err.Clear

End Function


Public Function Pillar_Format_Nearest( _
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
' SIGNATURE
'   Pillar_Format_Nearest(DtStart, DtEnd) -> String
'
' INPUTS
'   DtStart
'     Already-parsed start date. A time component is tolerated and discarded.
'
'   DtEnd
'     Already-parsed end date. A time component is tolerated and discarded.
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
'       * build the nearest whole-week anchor, eligible only up to 3W
'       * build the nearest calendar-month anchor from the start date
'       * measure both against the actual end date and emit the closer one
'       * ties resolve to the month-style label
'
' ROUNDING POLICY
'   - Weeks: nearest whole week on calendar days, half up, capped at 3W.
'   - Months: compare the floor and ceiling month anchors, take the closer,
'     ties round up. Anchors are built with DateAdd so month lengths are real.
'   - Family choice: whichever eligible anchor sits closer to the true end
'     date.
'
' ERROR POLICY
'   None. The caller owns parsing and worksheet messaging.
'
' DEPENDENCIES
'   None.
'
' NOTES
'   - Deliberately coarse: no mixed tokens like "1M2W". A pillar is a curve
'     bucket label, not a precise interval.
'   - Day tokens are reserved for intervals under a week so longer intervals
'     never collapse into raw day counts.
'   - The week family is capped at 3W for the same reason. A week anchor
'     always lands within 3 days of any target while a month anchor can sit
'     up to about 15 days away, so on pure distance the week family would win
'     almost every long interval: ten years and two weeks formatted as "524W"
'     before the cap. Market convention runs weeks only at the short end.
'   - The cap makes 4W and longer week labels unreachable. The 3W to 1M
'     boundary falls at 25 days, measured on the host: 24 days emits "3W" and
'     25 days emits "1M". Convention would more often call 25 to 27 days 3W
'     or 4W, so labels in that band sit one pillar long. The boundary is left
'     where the rounding rule puts it rather than special-cased, because a
'     boundary that cannot be derived from the rule is the kind that drifts.
'   - This produces a nearest tenor LABEL, not the nearest quoted pillar. Any
'     integer month can be emitted, including 7M and other months that are not
'     normally quoted, and no ceiling is applied to the year family. Snapping
'     to a pillar set is a different guarantee and belongs elsewhere.
'   - Anchors that would leave the supported window are treated as ineligible
'     rather than trapped. DateAdd raises error 5 there, and this routine
'     states that it does not raise; eligibility keeps that true without an
'     error handler reinterpreting a raise as a tenor.
'   - Eligibility is positional, so a label can depend on where in the
'     calendar the interval sits. An 11-day interval emits "2W" anywhere in
'     the range except the final fortnight of year 9999, where the 2W anchor
'     falls outside the window and the month family takes it as "1M".
'   - Both inputs are floored to midnight before comparison. Without that,
'     a later clock time on the same calendar day sets the sign prefix while
'     the day count reports zero, emitting "-0D".
'   - Day distances are taken by serial subtraction. Every operand is an
'     integral date, so the result matches DateDiff exactly at lower cost.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SignPrefix      As String   '"-" when DtEnd precedes DtStart
    Dim D0              As Date     'Normalized earlier date, date-only
    Dim D1              As Date     'Normalized later date, date-only
    Dim Serial0         As Double   'Serial of D0, reused by every distance
    Dim Serial1         As Double   'Serial of D1, reused by every distance

    Dim TotalDays       As Long     'Calendar days between D0 and D1

    Dim nWeeks          As Long     'Nearest whole week count (half up), capped
    Dim WeekSerial      As Double   'Serial of the week anchor
    Dim DistWeek        As Long     'Day distance from D1 to the week anchor
    Dim WeekOK          As Boolean  'TRUE when the week anchor is a candidate

    Dim MonthIndex0     As Long     'Absolute 0-based month index of D0

    Dim FloorMonths     As Long     'Largest month count not overshooting D1
    Dim CeilMonths      As Long     'Next month count after FloorMonths
    Dim NearestMonths   As Long     'Chosen month count
    Dim DistMonth       As Long     'Day distance from D1 to the chosen month anchor
    Dim DistFloor       As Long     'Day distance to the floor month anchor
    Dim DistCeil        As Long     'Day distance to the ceiling month anchor
    Dim CeilOK          As Boolean  'TRUE when the ceiling anchor stays in range

    Dim nYears          As Long     'Year component of NearestMonths
    Dim nMonths         As Long     'Residual month component of NearestMonths
    Dim OutTok          As String   'Unsigned output token

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Floor both inputs so the sign and the day count agree
        If Int(CDbl(DtEnd)) >= Int(CDbl(DtStart)) Then
            SignPrefix = vbNullString
            D0 = Int(CDbl(DtStart))
            D1 = Int(CDbl(DtEnd))
        Else
            SignPrefix = "-"
            D0 = Int(CDbl(DtEnd))
            D1 = Int(CDbl(DtStart))
        End If

    'Cache both serials; every distance below is a subtraction on these
        Serial0 = CDbl(D0)
        Serial1 = CDbl(D1)
    'Measure the interval once
        TotalDays = CLng(Serial1 - Serial0)

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

    'The week family runs only to 3W; beyond that months own the label
        If nWeeks >= 1 And nWeeks <= 3 Then
            'The anchor may round past D1 and out of the supported window
                WeekSerial = Serial0 + 7# * CDbl(nWeeks)
                If WeekSerial <= CDbl(KPR_MAX_DATE) Then
                    DistWeek = Abs(CLng(WeekSerial - Serial1))
                    WeekOK = True
                End If
        End If

'------------------------------------------------------------------------------
' NEAREST MONTH ANCHOR
'------------------------------------------------------------------------------
    'Month index of the start date, needed only past the short-interval exits
        MonthIndex0 = Year(D0) * 12& + Month(D0) - 1&

    'Crossed month boundaries, then step back if adding them overshoots
        FloorMonths = DateDiff("m", D0, D1)
        If FloorMonths > 0 Then
            If DateAdd("m", FloorMonths, D0) > D1 Then FloorMonths = FloorMonths - 1
        End If

    'Defensive clamp
        If FloorMonths < 0 Then FloorMonths = 0

    'Below one month the only meaningful month tenor is 1M
        If FloorMonths = 0 Then
            'The tenor is 1M whether or not its anchor can be constructed
                NearestMonths = 1
            'The 1M anchor may step past the end of the supported window
                If (MonthIndex0 + 1&) <= KPR_MAX_MONTHIDX Then
                    DistMonth = Abs(CLng(CDbl(DateAdd("m", 1, D0)) - Serial1))
                Else
                    'No anchor exists to measure; approximate on a nominal month
                        DistMonth = Abs(TotalDays - 30&)
                End If
        Else
            'The ceiling anchor may step past the end of the supported window
                CeilMonths = FloorMonths + 1
                CeilOK = ((MonthIndex0 + CeilMonths) <= KPR_MAX_MONTHIDX)
            'Measure the floor anchor, which is always in range
                DistFloor = Abs(CLng(CDbl(DateAdd("m", FloorMonths, D0)) - Serial1))
            'Take the closer anchor; ties round up
                If CeilOK Then
                    DistCeil = Abs(CLng(CDbl(DateAdd("m", CeilMonths, D0)) - Serial1))
                    If DistCeil <= DistFloor Then
                        NearestMonths = CeilMonths
                        DistMonth = DistCeil
                    Else
                        NearestMonths = FloorMonths
                        DistMonth = DistFloor
                    End If
                Else
                    NearestMonths = FloorMonths
                    DistMonth = DistFloor
                End If
        End If

'------------------------------------------------------------------------------
' CHOOSE TENOR FAMILY
'------------------------------------------------------------------------------
    'A strictly closer week anchor wins outright, when it is a candidate
        If WeekOK Then
            If DistWeek < DistMonth Then
                Pillar_Format_Nearest = SignPrefix & CStr(nWeeks) & "W"
                Exit Function
            End If
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
    'Defensive fallback; unreachable because NearestMonths is always >= 1
        If Len(OutTok) = 0 Then OutTok = "0D"

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Apply the direction prefix and return
        Pillar_Format_Nearest = SignPrefix & OutTok

End Function
