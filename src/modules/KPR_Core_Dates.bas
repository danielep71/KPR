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
'   - EndOfMonth_Core
'   - IsLeapYear_Core
'   - TryAddMonths_Core
'   - TryPillar_Parse
'   - Pillar_Format_Nearest
'
' VISIBILITY
'   Option Private Module. Members are Public only because VBA requires it for
'   cross-module calls inside the project. Nothing here is supported API.
'
' ALLOWED DEPENDENCIES
'   None. This module computes on already-validated inputs and never parses,
'   never inspects shape, and never constructs a worksheet error.
'
' SUPPORTED WINDOW
'       KPR_MIN_DATE = 1900-03-01
'       KPR_MAX_DATE = 9999-12-31
'
'   The lower bound is serial 61 by design: it excludes the Excel 1900-system
'   region where worksheet serials and VBA dates disagree because of the
'   fictitious 29-Feb-1900.
'
'   The baseline applied this gate inside its combined parser. The gate is now
'   applied by the facade immediately after parsing, using the predicate below.
'   It remains a single gate on every accepted path; only its location moved,
'   so that KPR_Core_Parse need not depend on this module.
'
' SCOPE BOUNDARY (THIS REVISION)
'   Pillar rounding is NEAREST only, as migrated. FLOOR and CEILING arrive with
'   Opt_Rounding in issue #14. Duplicate pillar units still accumulate here;
'   the contract rejects them, and issue #15 makes that change.
'
' NOTES
'   - Date math follows Excel / VBA DateSerial, DateAdd, DateDiff and Weekday
'     semantics. Gregorian leap-year logic is used throughout.
'   - AddYears delegates to the single month shifter so year-roll edge cases
'     (29-Feb, short months, EOM) cannot drift between the two functions.
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

Public Function EndOfMonth_Core( _
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

Public Function IsLeapYear_Core( _
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

Public Function TryAddMonths_Core( _
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
