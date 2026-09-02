Attribute VB_Name = "KPR_REGRESSION_TESTS"
'==============================================================================
' MODULE: KPR_REGRESSION_TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Deterministic regression suites for the KPR date layer, launched from one
'   entry point.
'
'   This revision covers the scalar input contract implemented in
'   KPR_Core_Parse: every accepted and rejected scalar input class named in
'   sections 3.1, 3.2, 3.3 and 7 of the date-layer contract.
'
' SUITE REGISTRY
'   KPR_Tests_RunAll dispatches to the suites below. Passing a suite name runs
'   only that suite; omitting the argument runs all of them in order.
'
'       date-type     accepted and rejected scalar types at a date position
'       date-text     ISO acceptance, locale and numeric text, impossible dates
'       date-window   the supported window and the values adjacent to it
'       integer       fractions, Boolean, text, and range before integrality
'       control       optional Boolean controls and their defaults
'       boundary      window mapping, propagation and composition, via the facade
'       mapper        ErrForCondition mappings and its refusals
'       host          the date-system guard and diagnostic from direct VBA,
'                     and #N/A provenance
'       pillar        rounding modes, the derived 3W / 1M boundary, the week
'                     cap, grammar conditions and format/parse consistency
'       surface       the five members added by #15, the YearIn conversions,
'                     the singular pillar name and the window floor boundaries
'
'   Later issues add suites by writing one Private Sub and one Case line. The
'   dispatcher is deliberately the only shared machinery.
'
' SCOPE
'   - Condition classification is asserted directly against KPR_Core_Parse, so a
'     test names the exact contract condition rather than only the error value.
'   - Window, propagation, wrapper and composition behaviour is asserted through
'     KPR_DATES_DAYS, because those decisions belong to the public boundary.
'   - Nothing here is a milestone fixture. Issue #19 owns independently
'     generated fixtures and issue #20 owns the runner and evidence interface.
'     This module deliberately does not anticipate either: it reports a count
'     and a list of failures, and defines no evidence format.
'
' USAGE
'   Three entry points, all reaching the same implementation:
'
'       KPR_Tests_Run                 every pure suite; appears in the Alt+F8
'                                     macro list and prints its report to the
'                                     Immediate window
'       KPR_Tests_RunSuite "integer"  one pure suite, printed the same way
'       KPR_Tests_RunAll()            returns a two-column array for a worksheet
'                                     or for programmatic use
'
'   And one stateful entry point that is deliberately NOT reachable from the
'   dispatcher above:
'
'       KPR_Tests_RunHost             creates a scratch workbook, exercises the
'                                     real worksheet-Range caller path under
'                                     1900 and 1904, and closes the workbook.
'                                     Macro-only: it adds and closes a
'                                     workbook, which cannot legally happen
'                                     inside a worksheet function call.
'
'   KPR_Tests_RunAll returns an array, so it cannot be inspected with ? in the
'   Immediate window and cannot appear in the macro list. Use KPR_Tests_Run
'   there. On a worksheet, enter =KPR_Tests_RunAll() and let it spill.
'
'   Output is Debug.Print only. No MsgBox, no worksheet selection, no
'   ActiveWorkbook, no UI of any kind.
'
' CONDITION IDENTIFIERS
'   Assertions carry the semantic identifier string from the contract registry,
'   never the internal enum number. The enum numbering is an implementation
'   detail and must not reach fixtures or evidence.
'
' ALLOWED DEPENDENCIES
'   KPR_Core_Parse, KPR_Core_Dates and KPR_Core_Err, called directly to assert
'   exact condition classification, and KPR_DATES_DAYS for boundary behaviour. No other core is
'   reachable from here. The stateful host runner additionally uses
'   Excel.Workbooks.Add and the scratch workbook it creates, always through the
'   exact object reference and never through ActiveWorkbook.
'
' UPDATED
'   2026-09-02
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit variable declarations

    'This module deliberately does NOT declare Option Private Module. A harness
    'that cannot be launched is not a harness: Option Private Module hides its
    'members from the Alt+F8 macro list, from Application.Run and from the
    'worksheet, leaving the Immediate window as the only entry. Tests are
    'developer-facing, they declare no KPR_Dates_* name, and the static gate
    'constrains visibility only for core and facade modules.

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Accumulated failures and assertion count for the current run
        Private mFailures   As Collection
        Private mChecks     As Long


'
'------------------------------------------------------------------------------
'
'                              MACRO ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub KPR_Tests_Run()
'
'==============================================================================
'                                 KPR_Tests_Run
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every suite and writes the report to the Immediate window.
'
'   This is the entry point to use from the Alt+F8 macro list or by typing
'   KPR_Tests_Run in the Immediate window. It takes no arguments and returns
'   nothing, which is what makes it visible there; KPR_Tests_RunAll returns an
'   array and therefore cannot appear in the macro list.
'
' OUTPUT
'   Debug.Print only. No MsgBox, no worksheet write, no selection, no
'   ActiveWorkbook access.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Every suite in registry order
        ReportRun "all"

End Sub

Public Sub KPR_Tests_RunSuite( _
    ByVal SuiteName As String)
'
'==============================================================================
'                              KPR_Tests_RunSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Runs one named suite and writes the report to the Immediate window.
'
' USAGE
'   KPR_Tests_RunSuite "date-text"
'
' NOTES
'   - Takes an argument, so it does not appear in the Alt+F8 list. Call it from
'     the Immediate window.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'One suite from the registry
        ReportRun SuiteName

End Sub

Private Sub ReportRun( _
    ByVal SuiteName As String)
'
'==============================================================================
'                                   ReportRun
'------------------------------------------------------------------------------
' PURPOSE
'   Runs a suite through the same function the worksheet uses, then prints the
'   result so a run is never silent.
'
' NOTES
'   - Deliberately shares KPR_Tests_RunAll rather than duplicating dispatch, so
'     the macro path and the worksheet path cannot report different outcomes.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Results     As Variant      'Result of the run
    Dim I           As Long         'Row cursor

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Reuse the single implementation
        Results = KPR_Tests_RunAll(SuiteName)

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'A refused suite name comes back as a plain string
        If Not IsArray(Results) Then
            Debug.Print "KPR regression: " & CStr(Results)
            Exit Sub
        End If

    'Summary first, so a passing run is one visible line
        Debug.Print "KPR regression [" & SuiteName & "]  " & _
                    CStr(Results(1, 1)) & "  " & CStr(Results(1, 2))

    'Then one line per failure, if any
        For I = 2 To UBound(Results, 1)
            Debug.Print "  FAIL  " & CStr(Results(I, 1)) & " : " & CStr(Results(I, 2))
        Next I

End Sub

'
'------------------------------------------------------------------------------
'
'                                  ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Function KPR_Tests_RunAll( _
    Optional ByVal Opt_Suite As Variant) _
    As Variant
'
'==============================================================================
'                                KPR_Tests_RunAll
'------------------------------------------------------------------------------
' PURPOSE
'   Runs one suite or every suite and returns the result.
'
' INPUTS
'   Opt_Suite
'     Omitted or Empty runs every suite. Otherwise a suite name from the
'     registry in the module header, compared case-insensitively.
'
' RETURNS
'   Variant
'     2-D array. Row 1 is the assertion count and the failure count. Each
'     further row is a failing case label and its detail.
'
'     An unknown suite name returns a single explanatory string rather than an
'     empty pass, so a typo cannot look like success.
'
' ERROR POLICY
'   - Never raises to the caller and never shows UI.
'   - An unexpected runtime error is reported as a failure, because reaching a
'     defensive handler is itself a defect.
'
' NOTES
'   - The suite list is duplicated in neither direction: RunSuite owns dispatch
'     and the header documents it. Adding a suite touches one Case line.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SuiteName   As String       'Requested suite, normalized
    Dim Results     As Variant      'Returned summary and failure detail
    Dim I           As Long         'Row cursor
    Dim OuterFails  As Collection   'State of any run this call interrupted
    Dim OuterChecks As Long         'Its assertion count

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain any unexpected error as a reported failure
        On Error GoTo Err_Handler

    'This function is worksheet-callable, so a cell holding =KPR_Tests_RunAll()
    'can be recalculated in the MIDDLE of another run: the stateful host
    'runner calls Application.Calculate, and that reaches such a cell. Save
    'whatever run is in progress and restore it on exit, so an interrupted
    'runner reports its own results rather than this call's.
        Set OuterFails = mFailures
        OuterChecks = mChecks

    'Fresh state for this run
        Set mFailures = New Collection
        mChecks = 0

    'An omitted or blank suite means every suite
        If IsMissing(Opt_Suite) Then
            SuiteName = "all"
        ElseIf IsEmpty(Opt_Suite) Then
            SuiteName = "all"
        Else
            SuiteName = LCase$(Trim$(CStr(Opt_Suite)))
        End If

'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    'Run the requested suite, or refuse a name that is not in the registry
        If Not RunSuite(SuiteName) Then
            KPR_Tests_RunAll = "unknown suite: " & SuiteName
            Set mFailures = OuterFails
            mChecks = OuterChecks
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Summary row plus one row per failure
        ReDim Results(1 To mFailures.Count + 1, 1 To 2)
        Results(1, 1) = "checks: " & CStr(mChecks)
        Results(1, 2) = "failures: " & CStr(mFailures.Count)

    'Detail rows
        For I = 1 To mFailures.Count
            Results(I + 1, 1) = mFailures(I)(0)
            Results(I + 1, 2) = mFailures(I)(1)
        Next I

    'Return the array, then hand back any interrupted run's state
        KPR_Tests_RunAll = Results
        Set mFailures = OuterFails
        mChecks = OuterChecks
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'A raise reaching here is a defect, so report it rather than swallowing it
        KPR_Tests_RunAll = "unexpected runtime error " & CStr(Err.Number) & ": " & Err.Description
        Set mFailures = OuterFails
        mChecks = OuterChecks

End Function

Private Function RunSuite( _
    ByVal SuiteName As String) _
    As Boolean
'
'==============================================================================
'                                   RunSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Maps a suite name to its cases. The single place that knows what suites
'   exist.
'
' RETURNS
'   Boolean
'     TRUE  => the name was recognized and its cases ran
'     FALSE => the name is not in the registry; nothing ran
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Unrecognized until proven otherwise, so a typo cannot pass silently
        RunSuite = True

'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    Select Case SuiteName

        Case "all"
            'Every suite, in registry order
                Run_DateTypeCases
                Run_DateTextCases
                Run_DateWindowCases
                Run_IntegerCases
                Run_ControlCases
                Run_BoundaryCases
                Run_MapperCases
                Run_HostCases
                Run_PillarCases
                Run_SurfaceCases

        Case "date-type":       Run_DateTypeCases
        Case "date-text":       Run_DateTextCases
        Case "date-window":     Run_DateWindowCases
        Case "integer":         Run_IntegerCases
        Case "control":         Run_ControlCases
        Case "boundary":        Run_BoundaryCases
        Case "mapper":          Run_MapperCases
        Case "host":            Run_HostCases
        Case "pillar":          Run_PillarCases
        Case "surface":         Run_SurfaceCases

        Case Else
            'Not in the registry
                RunSuite = False

    End Select

End Function

'
'------------------------------------------------------------------------------
'
'                       SUITE - HOST POLICY FROM DIRECT VBA
'
'------------------------------------------------------------------------------
'

Private Sub Run_HostCases()
'
' The date-system guard and diagnostic as seen from direct VBA, plus #N/A
' provenance. This suite runs without a worksheet caller, so it can only
' assert the "no worksheet host identified" path and the value identity of the
' two #N/A sources. The worksheet-Range paths are exercised by the stateful
' KPR_Tests_RunHost macro.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const ERR_NA    As Long = 2042      'Excel #N/A error number
    Dim Reported    As Variant          'HostDateSystem result
    Dim LibraryNA   As Variant          'Library-produced #N/A
    Dim IncomingNA  As Variant          'Propagated incoming #N/A

'------------------------------------------------------------------------------
' DIAGNOSTIC UNDER DIRECT VBA
'------------------------------------------------------------------------------
    'No worksheet host can be identified from VBA, so the documented answer is
    '1900 rather than an error
        mChecks = mChecks + 1
        Reported = KPR_Dates_HostDateSystem()
        If VarType(Reported) = vbError Then
            Record "host/diagnostic direct vba", "expected 1900, got Excel error " & CStr(CLng(Reported))
        ElseIf CLng(Reported) <> 1900 Then
            Record "host/diagnostic direct vba", "expected 1900, got " & CStr(Reported)
        End If

'------------------------------------------------------------------------------
' GUARD PASSES UNDER DIRECT VBA
'------------------------------------------------------------------------------
    'Every value function must proceed on the certified path, not refuse
        AssertDateResult "host/guard passes DayOfWeek path", _
                         KPR_Dates_BeginOfMonth("2026-03-15"), DateSerial(2026, 3, 1)
        AssertDateResult "host/guard passes arithmetic path", _
                         KPR_Dates_AddWeeks("2026-03-15", 1), DateSerial(2026, 3, 22)
        AssertDateResult "host/guard passes locator path", _
                         KPR_Dates_NthWeekdayOfMonth(2026, 3, 1, 1), DateSerial(2026, 3, 2)

'------------------------------------------------------------------------------
' #N/A PROVENANCE
'------------------------------------------------------------------------------
    'Both host conditions map to #N/A
        AssertErrorValue "host/map DATE1904", ErrForCondition(KPR_COND_HOST_DATE1904), ERR_NA
        AssertErrorValue "host/map UNRESOLVED", ErrForCondition(KPR_COND_HOST_UNRESOLVED), ERR_NA

    'A library #N/A and a propagated incoming #N/A are the SAME Excel value.
    'The two cases keep separate labels precisely because the value cannot tell
    'them apart; only the diagnostic can.
        LibraryNA = ErrForCondition(KPR_COND_HOST_DATE1904)
        IncomingNA = KPR_Dates_DayOfWeek(CVErr(xlErrNA))
        AssertErrorValue "provenance/library NA is NA", LibraryNA, ERR_NA
        AssertErrorValue "provenance/propagated NA is NA", IncomingNA, ERR_NA
        mChecks = mChecks + 1
        If CLng(LibraryNA) <> CLng(IncomingNA) Then
            Record "provenance/values identical", "library and propagated #N/A differ as values"
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                       SUITE - PILLAR ROUNDING AND GRAMMAR
'
'------------------------------------------------------------------------------
'

Private Sub Run_PillarCases()
'
' The three rounding modes over one uniform candidate set, the derived
' 3W / 1M boundary, the week cap under every mode, the CEILING gap, grammar
' conditions by identifier, and format/parse consistency.
'
' Expected tokens were derived from the contract's rule independently of the
' VBA implementation, so a passing case checks the implementation against the
' rule rather than against itself.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const ERR_VALUE As Long = 2015      'Excel #VALUE! error number
    Const ERR_NUM   As Long = 2036      'Excel #NUM! error number
    Const ERR_NA    As Long = 2042      'Excel #N/A error number
    Dim S           As Date             'Reference start date
    Dim D           As Long             'Interval cursor for the consistency loop
    Dim Token       As Variant          'Emitted token
    Dim Back        As Variant          'Re-parsed anchor
    Dim Again       As Variant          'Token re-emitted from the anchor
    Dim ModeName    As Variant          'Mode cursor
    Dim TokenOut    As String           'Core formatter output
    Dim Cond        As KPR_Condition    'Core condition

'------------------------------------------------------------------------------
' EXACT DAY PILLARS UNDER EVERY MODE
'------------------------------------------------------------------------------
    S = DateSerial(2026, 3, 15)
        AssertPillar "pillar/0D", S, S, "NEAREST", "0D"
        AssertPillar "pillar/5D nearest", S, S + 5, "NEAREST", "5D"
        AssertPillar "pillar/5D floor", S, S + 5, "FLOOR", "5D"
        AssertPillar "pillar/5D ceiling", S, S + 5, "CEILING", "5D"
        AssertPillar "pillar/-5D", S, S - 5, "NEAREST", "-5D"

'------------------------------------------------------------------------------
' DERIVED 3W / 1M BOUNDARY UNDER NEAREST
'------------------------------------------------------------------------------
    'From Mar 15 the 1M anchor is Apr 15, 31 days out: 25 stays 3W, 26 is 1M
        AssertPillar "boundary/24d Mar15", S, S + 24, "NEAREST", "3W"
        AssertPillar "boundary/25d Mar15", S, S + 25, "NEAREST", "3W"
        AssertPillar "boundary/26d Mar15", S, S + 26, "NEAREST", "1M"
        AssertPillar "boundary/27d Mar15", S, S + 27, "NEAREST", "1M"

    'From Jan 31 the 1M anchor clips to Feb 28, 28 days out: 25 is already 1M
        AssertPillar "boundary/24d Jan31", DateSerial(2026, 1, 31), DateSerial(2026, 2, 24), "NEAREST", "3W"
        AssertPillar "boundary/25d Jan31", DateSerial(2026, 1, 31), DateSerial(2026, 2, 25), "NEAREST", "1M"

'------------------------------------------------------------------------------
' THE THREE MODES ON ONE INTERVAL
'------------------------------------------------------------------------------
    'Twenty-five days: floor keeps the last week, ceiling reaches for the month
        AssertPillar "modes/25d floor", S, S + 25, "FLOOR", "3W"
        AssertPillar "modes/25d nearest", S, S + 25, "NEAREST", "3W"
        AssertPillar "modes/25d ceiling", S, S + 25, "CEILING", "1M"

    'Thirty days in a 31-day month: the week cap keeps FLOOR at 3W
        AssertPillar "modes/30d floor", S, S + 30, "FLOOR", "3W"
        AssertPillar "modes/30d nearest", S, S + 30, "NEAREST", "1M"
        AssertPillar "modes/30d ceiling", S, S + 30, "CEILING", "1M"

    'One hundred days: 3M is 92, 4M is 122
        AssertPillar "modes/100d floor", S, S + 100, "FLOOR", "3M"
        AssertPillar "modes/100d nearest", S, S + 100, "NEAREST", "3M"
        AssertPillar "modes/100d ceiling", S, S + 100, "CEILING", "4M"

'------------------------------------------------------------------------------
' EXACT ANCHORS ARE EXACT UNDER EVERY MODE
'------------------------------------------------------------------------------
        AssertPillar "exact/2W floor", S, S + 14, "FLOOR", "2W"
        AssertPillar "exact/2W nearest", S, S + 14, "NEAREST", "2W"
        AssertPillar "exact/2W ceiling", S, S + 14, "CEILING", "2W"
        AssertPillar "exact/1M floor", S, DateSerial(2026, 4, 15), "FLOOR", "1M"
        AssertPillar "exact/1M ceiling", S, DateSerial(2026, 4, 15), "CEILING", "1M"

'------------------------------------------------------------------------------
' NEGATIVE INTERVALS: ANCHORS FROM THE ORIGINAL START, SIGNED COUNTS
'------------------------------------------------------------------------------
    'Backwards from Mar 15 the 1M anchor is Feb 15, 28 days: nearest at 25 is 1M
        AssertPillar "negative/25d floor", S, S - 25, "FLOOR", "-3W"
        AssertPillar "negative/25d nearest", S, S - 25, "NEAREST", "-1M"
        AssertPillar "negative/25d ceiling", S, S - 25, "CEILING", "-1M"

'------------------------------------------------------------------------------
' TIE RULES
'------------------------------------------------------------------------------
    'Forty-six days: 1M (31) and 2M (61) are both 15 away; the larger count wins
        AssertPillar "tie/month-month nearest", S, S + 46, "NEAREST", "2M"
        AssertPillar "tie/month-month floor", S, S + 46, "FLOOR", "1M"
        AssertPillar "tie/month-month ceiling", S, S + 46, "CEILING", "2M"

'------------------------------------------------------------------------------
' WEEK CAP UNDER FLOOR AND CEILING
'------------------------------------------------------------------------------
        AssertPillar "cap/20d floor", S, S + 20, "FLOOR", "2W"
        AssertPillar "cap/20d ceiling", S, S + 20, "CEILING", "3W"
        AssertPillar "cap/28d floor", S, S + 28, "FLOOR", "3W"
        AssertPillar "cap/28d ceiling", S, S + 28, "CEILING", "1M"

'------------------------------------------------------------------------------
' MONTH-END AND LEAP-DAY ANCHORS
'------------------------------------------------------------------------------
        AssertPillar "leap/Jan31 to Feb29 floor", DateSerial(2024, 1, 31), DateSerial(2024, 2, 29), "FLOOR", "1M"
        AssertPillar "leap/Jan31 to Feb29 ceiling", DateSerial(2024, 1, 31), DateSerial(2024, 2, 29), "CEILING", "1M"
        AssertPillar "eom/Jan31 to Mar31 nearest", DateSerial(2026, 1, 31), DateSerial(2026, 3, 31), "NEAREST", "2M"
        AssertPillar "eom/Jan31 to Mar31 floor", DateSerial(2026, 1, 31), DateSerial(2026, 3, 31), "FLOOR", "2M"

'------------------------------------------------------------------------------
' YEARS
'------------------------------------------------------------------------------
        AssertPillar "years/12M", S, DateSerial(2027, 3, 15), "NEAREST", "1Y"
        AssertPillar "years/13M", S, DateSerial(2027, 4, 15), "NEAREST", "1Y1M"

'------------------------------------------------------------------------------
' OUT-OF-WINDOW ANCHORS ARE EXCLUDED, NEVER APPROXIMATED
'------------------------------------------------------------------------------
    'Final fortnight of 9999: 2W and 1M leave the window; 1W is the only anchor
        AssertPillar "window/final fortnight nearest", DateSerial(9999, 12, 20), DateSerial(9999, 12, 31), "NEAREST", "1W"
        AssertPillar "window/final fortnight floor", DateSerial(9999, 12, 20), DateSerial(9999, 12, 31), "FLOOR", "1W"

    'CEILING with no in-window candidate that reaches the end is #NUM!
        AssertErrorValue "window/ceiling gap facade", _
                         KPR_Dates_PillarFromDates(DateSerial(9999, 12, 15), DateSerial(9999, 12, 31), "CEILING"), ERR_NUM
        mChecks = mChecks + 1
        If TryPillar_Format(DateSerial(9999, 12, 15), DateSerial(9999, 12, 31), KPR_ROUND_CEILING, TokenOut, Cond) Then
            Record "window/ceiling gap core", "expected failure, got " & TokenOut
        ElseIf ConditionName(Cond) <> "RESULT_WINDOW" Then
            Record "window/ceiling gap core", "expected RESULT_WINDOW got " & ConditionName(Cond)
        End If

'------------------------------------------------------------------------------
' OPT_ROUNDING CONTROL
'------------------------------------------------------------------------------
    'Omitted, Empty, padded and mixed-case forms all resolve
        mChecks = mChecks + 1
        If KPR_Dates_PillarFromDates(S, S + 25) <> KPR_Dates_PillarFromDates(S, S + 25, "NEAREST") Then
            Record "control/omitted equals NEAREST", "omitted and explicit disagree"
        End If
        AssertPillar "control/empty", S, S + 25, Empty, "3W"
        AssertPillar "control/padded floor", S, S + 25, "  floor" & vbTab, "3W"
        AssertPillar "control/mixed case ceiling", S, S + 25, "Ceiling", "1M"

    'Rejections
        AssertErrorValue "control/unknown token", KPR_Dates_PillarFromDates(S, S + 25, "ROUND"), ERR_VALUE
        AssertErrorValue "control/numeric", KPR_Dates_PillarFromDates(S, S + 25, 1), ERR_VALUE
        AssertErrorValue "control/boolean", KPR_Dates_PillarFromDates(S, S + 25, True), ERR_VALUE
        AssertErrorValue "control/propagates", KPR_Dates_PillarFromDates(S, S + 25, CVErr(xlErrNA)), ERR_NA

'------------------------------------------------------------------------------
' GRAMMAR CONDITIONS BY IDENTIFIER
'------------------------------------------------------------------------------
        AssertPillarParse "grammar/duplicate unit", "1M2M", "PILLAR_DUPLICATE_UNIT"
        AssertPillarParse "grammar/duplicate day", "3D4D", "PILLAR_DUPLICATE_UNIT"
        AssertPillarParse "grammar/signed alias", "-ON", "PILLAR_ALIAS_SIGNED"
        AssertPillarParse "grammar/plus alias", "+T/N", "PILLAR_ALIAS_SIGNED"
        AssertPillarParse "grammar/unknown unit", "1X", "PILLAR_TOKEN_MALFORMED"
        AssertPillarParse "grammar/internal space", "1 M", "PILLAR_TOKEN_MALFORMED"
        AssertPillarParse "grammar/empty", "", "PILLAR_TOKEN_MALFORMED"
        AssertPillarParse "grammar/numeric payload", 12, "PILLAR_TYPE_REJECTED"
        AssertPillarParse "grammar/any order", "3d2w", "NONE"
        AssertPillarParse "grammar/alias", "on", "NONE"

    'The facade maps them, and an incoming error at the Pillar slot propagates
        AssertErrorValue "grammar/facade duplicate", KPR_Dates_DateFromPillar(S, "1M2M"), ERR_VALUE
        AssertErrorValue "grammar/facade signed alias", KPR_Dates_DateFromPillar(S, "-ON"), ERR_VALUE
        AssertErrorValue "grammar/facade propagates", KPR_Dates_DateFromPillar(S, CVErr(xlErrNA)), ERR_NA

'------------------------------------------------------------------------------
' NON-INVARIANT ROUND TRIPS, STATED
'------------------------------------------------------------------------------
    'A rounded token re-parses to its anchor, not to the original end date
        AssertDateResult "roundtrip/25d nearest lands on anchor", _
                         KPR_Dates_DateFromPillar(S, KPR_Dates_PillarFromDates(S, S + 25)), S + 21
        AssertDateResult "roundtrip/26d nearest lands on anchor", _
                         KPR_Dates_DateFromPillar(S, KPR_Dates_PillarFromDates(S, S + 26)), DateSerial(2026, 4, 15)
        AssertDateResult "roundtrip/exact 2W lands on end", _
                         KPR_Dates_DateFromPillar(S, KPR_Dates_PillarFromDates(S, S + 14)), S + 14

'------------------------------------------------------------------------------
' FORMAT / PARSE MUTUAL CONSISTENCY
'------------------------------------------------------------------------------
    'For every interval and mode, the emitted token re-parses to an anchor
    'that re-emits the same token. Idempotence is the precise statement of
    '"mutually consistent for every supported token".
        For Each ModeName In Array("NEAREST", "FLOOR", "CEILING")
            For D = 7 To 60
                mChecks = mChecks + 1
                Token = KPR_Dates_PillarFromDates(S, S + D, ModeName)
                If VarType(Token) = vbError Then
                    Record "consistency/" & ModeName & " " & CStr(D) & "d", "emit failed"
                Else
                    Back = KPR_Dates_DateFromPillar(S, Token)
                    If VarType(Back) = vbError Then
                        Record "consistency/" & ModeName & " " & CStr(D) & "d", "token " & Token & " does not parse"
                    Else
                        Again = KPR_Dates_PillarFromDates(S, Back, ModeName)
                        If CStr(Again) <> CStr(Token) Then
                            Record "consistency/" & ModeName & " " & CStr(D) & "d", _
                                   "token " & Token & " re-emits as " & CStr(Again)
                        End If
                    End If
                End If
            Next D
        Next ModeName

End Sub

Private Sub AssertPillar( _
    ByVal Label As String, _
    ByVal StartDate As Date, _
    ByVal EndDate As Date, _
    ByVal Mode As Variant, _
    ByVal ExpectToken As String)
'
' Asserts the exact token the facade emits under a mode.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Actual      As Variant      'Facade result

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Count every assertion
        mChecks = mChecks + 1

    'Emit under the requested mode
        Actual = KPR_Dates_PillarFromDates(StartDate, EndDate, Mode)

    'Compare as text
        If VarType(Actual) = vbError Then
            Record Label, "expected " & ExpectToken & ", got Excel error " & CStr(CLng(Actual))
        ElseIf CStr(Actual) <> ExpectToken Then
            Record Label, "expected " & ExpectToken & " got " & CStr(Actual)
        End If

End Sub

Private Sub AssertPillarParse( _
    ByVal Label As String, _
    ByVal PillarIn As Variant, _
    ByVal ExpectCondition As String)
'
' Asserts the exact condition identifier from the calendar core's parser.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Months      As Double           'Parsed month delta
    Dim Days        As Double           'Parsed day delta
    Dim Cond        As KPR_Condition    'Reported condition

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Count every assertion
        mChecks = mChecks + 1

    'Parse and compare the identifier, whatever the Boolean says
        TryPillar_Parse PillarIn, Months, Days, Cond
        If ConditionName(Cond) <> ExpectCondition Then
            Record Label, "expected " & ExpectCondition & " got " & ConditionName(Cond)
        End If

End Sub

'
'------------------------------------------------------------------------------
'
'                       SUITE - COMPLETED 22-NAME SURFACE
'
'------------------------------------------------------------------------------
'

Private Sub Run_SurfaceCases()
'
' The five members added by #15, the YearIn conversions of DaysInYear and
' IsLeapYear, the singular DateFromPillar name, and the window-floor
' boundaries where a correct calendar answer is outside the supported window.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const ERR_VALUE As Long = 2015      'Excel #VALUE! error number
    Const ERR_NUM   As Long = 2036      'Excel #NUM! error number
    Const ERR_NA    As Long = 2042      'Excel #N/A error number
    Dim R           As Variant          'Facade result

'------------------------------------------------------------------------------
' ADDDAYS
'------------------------------------------------------------------------------
        AssertDateResult "adddays/positive", KPR_Dates_AddDays("2026-03-15", 10), DateSerial(2026, 3, 25)
        AssertDateResult "adddays/negative", KPR_Dates_AddDays("2026-03-15", -15), DateSerial(2026, 2, 28)
        AssertDateResult "adddays/zero", KPR_Dates_AddDays("2026-03-15", 0), DateSerial(2026, 3, 15)
        AssertDateResult "adddays/leap crossing", KPR_Dates_AddDays("2024-02-28", 1), DateSerial(2024, 2, 29)
        AssertDateResult "adddays/year crossing", KPR_Dates_AddDays("2026-12-31", 1), DateSerial(2027, 1, 1)
        AssertErrorValue "adddays/past ceiling", KPR_Dates_AddDays("9999-12-31", 1), ERR_NUM
        AssertErrorValue "adddays/below floor", KPR_Dates_AddDays("1900-03-01", -1), ERR_NUM
        AssertErrorValue "adddays/fractional", KPR_Dates_AddDays("2026-03-15", 1.5), ERR_VALUE
        AssertErrorValue "adddays/propagates", KPR_Dates_AddDays("2026-03-15", CVErr(xlErrNA)), ERR_NA
        AssertDateResult "adddays/max shift in window", KPR_Dates_AddDays("1900-03-01", 2958404), DateSerial(9999, 12, 31)

'------------------------------------------------------------------------------
' QUARTER BOUNDARIES
'------------------------------------------------------------------------------
        AssertDateResult "quarter/begin Q1", KPR_Dates_BeginOfQuarter("2026-02-10"), DateSerial(2026, 1, 1)
        AssertDateResult "quarter/begin Q2", KPR_Dates_BeginOfQuarter("2026-05-31"), DateSerial(2026, 4, 1)
        AssertDateResult "quarter/begin Q3", KPR_Dates_BeginOfQuarter("2026-07-01"), DateSerial(2026, 7, 1)
        AssertDateResult "quarter/begin Q4", KPR_Dates_BeginOfQuarter("2026-12-31"), DateSerial(2026, 10, 1)
        AssertDateResult "quarter/end Q1 leap", KPR_Dates_EndOfQuarter("2024-01-15"), DateSerial(2024, 3, 31)
        AssertDateResult "quarter/end Q2", KPR_Dates_EndOfQuarter("2026-04-01"), DateSerial(2026, 6, 30)
        AssertDateResult "quarter/end Q3", KPR_Dates_EndOfQuarter("2026-09-30"), DateSerial(2026, 9, 30)
        AssertDateResult "quarter/end Q4 year edge", KPR_Dates_EndOfQuarter("2026-10-01"), DateSerial(2026, 12, 31)
        AssertDateResult "quarter/end at ceiling", KPR_Dates_EndOfQuarter("9999-11-11"), DateSerial(9999, 12, 31)

'------------------------------------------------------------------------------
' YEAR BOUNDARIES
'------------------------------------------------------------------------------
        AssertDateResult "year/begin", KPR_Dates_BeginOfYear("2026-07-04"), DateSerial(2026, 1, 1)
        AssertDateResult "year/begin leap", KPR_Dates_BeginOfYear("2024-12-31"), DateSerial(2024, 1, 1)
        AssertDateResult "year/end", KPR_Dates_EndOfYear("2026-01-01"), DateSerial(2026, 12, 31)
        AssertDateResult "year/end at ceiling", KPR_Dates_EndOfYear("9999-01-01"), DateSerial(9999, 12, 31)
        AssertDateResult "year/begin at floor year+1", KPR_Dates_BeginOfYear("1901-06-15"), DateSerial(1901, 1, 1)

'------------------------------------------------------------------------------
' WINDOW FLOOR: A CORRECT CALENDAR ANSWER OUTSIDE THE WINDOW IS #NUM!
'------------------------------------------------------------------------------
    'Both boundaries of Q1/1900 name 1900-01-01, which is outside the window
        AssertErrorValue "floor/BeginOfYear 1900", KPR_Dates_BeginOfYear("1900-03-01"), ERR_NUM
        AssertErrorValue "floor/BeginOfQuarter 1900", KPR_Dates_BeginOfQuarter("1900-03-01"), ERR_NUM
        AssertErrorValue "floor/BeginOfYear 1900-12-31", KPR_Dates_BeginOfYear("1900-12-31"), ERR_NUM

    'The rejection is result-specific: the quarter END of the same input is fine
        AssertDateResult "floor/EndOfQuarter 1900-03-01", KPR_Dates_EndOfQuarter("1900-03-01"), DateSerial(1900, 3, 31)
        AssertDateResult "floor/EndOfYear 1900", KPR_Dates_EndOfYear("1900-03-01"), DateSerial(1900, 12, 31)
        AssertDateResult "floor/BeginOfQuarter Q2 1900", KPR_Dates_BeginOfQuarter("1900-04-01"), DateSerial(1900, 4, 1)

'------------------------------------------------------------------------------
' YEARIN: DAYSINYEAR AND ISLEAPYEAR TAKE A CALENDAR YEAR
'------------------------------------------------------------------------------
        AssertLongResult "yearin/IsLeapYear 2024", KPR_Dates_IsLeapYear(2024), True
        AssertLongResult "yearin/IsLeapYear 2026", KPR_Dates_IsLeapYear(2026), False
        AssertLongResult "yearin/IsLeapYear 1900 century", KPR_Dates_IsLeapYear(1900), False
        AssertLongResult "yearin/IsLeapYear 2000 century", KPR_Dates_IsLeapYear(2000), True
        AssertLongResult "yearin/DaysInYear 2024", KPR_Dates_DaysInYear(2024), 366
        AssertLongResult "yearin/DaysInYear 1900", KPR_Dates_DaysInYear(1900), 365
        AssertLongResult "yearin/DaysInYear 9999", KPR_Dates_DaysInYear(9999), 365

    'A date is not a year: the old serial trap is now a type rejection
        AssertErrorValue "yearin/ISO text rejected", KPR_Dates_IsLeapYear("2024-01-01"), ERR_VALUE
        AssertErrorValue "yearin/date value rejected", KPR_Dates_DaysInYear(DateSerial(2024, 1, 1)), ERR_VALUE
        AssertErrorValue "yearin/fraction rejected", KPR_Dates_IsLeapYear(2024.5), ERR_VALUE

    'Domain is 1900 through 9999
        AssertErrorValue "yearin/1899", KPR_Dates_IsLeapYear(1899), ERR_VALUE
        AssertErrorValue "yearin/10000", KPR_Dates_DaysInYear(10000), ERR_VALUE
        AssertErrorValue "yearin/propagates", KPR_Dates_IsLeapYear(CVErr(xlErrNA)), ERR_NA

'------------------------------------------------------------------------------
' SINGULAR DATEFROMPILLAR
'------------------------------------------------------------------------------
        AssertDateResult "rename/DateFromPillar 1M", KPR_Dates_DateFromPillar("2026-01-31", "1M"), DateSerial(2026, 2, 28)
        AssertDateResult "rename/DateFromPillar -2W", KPR_Dates_DateFromPillar("2026-03-15", "-2W"), DateSerial(2026, 3, 1)

'------------------------------------------------------------------------------
' ELEMENT ORDER: FIRST FAILING ARGUMENT IN SIGNATURE ORDER
'------------------------------------------------------------------------------
    'Two bad arguments: the first one's error is the answer
        AssertErrorValue "order/date before count", KPR_Dates_AddDays("bad", CVErr(xlErrNA)), ERR_VALUE
        AssertErrorValue "order/count after good date", KPR_Dates_AddDays("2026-03-15", "bad"), ERR_VALUE

'------------------------------------------------------------------------------
' ARITHMETIC CLIPPING AND PRESERVATION (UNCHANGED, RE-PINNED AFTER THE SPLIT)
'------------------------------------------------------------------------------
        AssertDateResult "clip/AddMonths Jan31+1", KPR_Dates_AddMonths("2026-01-31", 1), DateSerial(2026, 2, 28)
        AssertDateResult "clip/AddMonths Apr30+1 clip", KPR_Dates_AddMonths("2026-04-30", 1), DateSerial(2026, 5, 30)
        AssertDateResult "clip/AddMonths Apr30+1 keepEOM", KPR_Dates_AddMonths("2026-04-30", 1, True), DateSerial(2026, 5, 31)
        AssertDateResult "clip/AddYears Feb29+1", KPR_Dates_AddYears("2024-02-29", 1), DateSerial(2025, 2, 28)
        AssertDateResult "clip/AddYears Feb29+4", KPR_Dates_AddYears("2024-02-29", 4), DateSerial(2028, 2, 29)
        AssertErrorValue "clip/AddYears overflow", KPR_Dates_AddYears("2026-03-15", 8000), ERR_NUM
        AssertDateResult "locator/5th Monday exists", KPR_Dates_NthWeekdayOfMonth(2026, 3, 1, 5), DateSerial(2026, 3, 30)
        AssertErrorValue "locator/5th Friday absent", KPR_Dates_NthWeekdayOfMonth(2026, 3, 5, 5), ERR_NUM
        AssertDateResult "locator/last Sunday", KPR_Dates_LastWeekdayOfMonth(2026, 3, 7), DateSerial(2026, 3, 29)

End Sub

Private Sub AssertLongResult( _
    ByVal Label As String, _
    ByVal Actual As Variant, _
    ByVal Expect As Variant)
'
' Asserts a scalar Long or Boolean result by value.
'

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Count every assertion
        mChecks = mChecks + 1

    'Compare only when the result is not an error
        If VarType(Actual) = vbError Then
            Record Label, "expected " & CStr(Expect) & ", got Excel error " & CStr(CLng(Actual))
        ElseIf CStr(Actual) <> CStr(Expect) Then
            Record Label, "expected " & CStr(Expect) & " got " & CStr(Actual)
        End If

End Sub

'
'------------------------------------------------------------------------------
'
'                    STATEFUL HOST RUNNER (MACRO ONLY, NOT DISPATCHED)
'
'------------------------------------------------------------------------------
'

Public Sub KPR_Tests_RunHost()
'
'==============================================================================
'                               KPR_Tests_RunHost
'------------------------------------------------------------------------------
' PURPOSE
'   Exercises the real worksheet-Range caller path under both date systems by
'   writing formulas into a scratch workbook, and prints the report to the
'   Immediate window.
'
' WHY THIS IS SEPARATE
'   KPR_Tests_RunAll is callable from a worksheet cell. Adding a workbook,
'   writing formulas and closing a workbook cannot legally happen inside a
'   worksheet function call, so this runner is a macro and is deliberately
'   not part of the dispatcher. It never joins "all".
'
' WHAT IT PROVES
'   - HostDateSystem reports 1900 in a 1900 workbook, 1904 after the same
'     workbook is toggled, and 1900 again after it is toggled back, each time
'     after ordinary calculation rather than a full rebuild.
'   - Value functions proceed under 1900 and return call-level #N/A under 1904.
'   - Host-produced and propagated #N/A remain separately labelled cases.
'
' WHAT IT DOES NOT PROVE
'   The Immediate-window probe and the non-Range host contexts are recorded in
'   Windows Excel by hand; see #13 and #29.
'
' STATE
'   - Creates one workbook through Workbooks.Add and holds the exact object.
'   - Every reference is through that object. ActiveWorkbook is never read.
'   - Formulas are qualified with the source workbook name so the UDFs resolve
'     regardless of which workbook is active.
'   - Calculation is switched to manual for the duration of the run and the
'     caller's mode is restored on every exit path. This is not cosmetic:
'     toggling Date1904 in automatic mode triggers a recalculation DURING the
'     property set, the volatile diagnostic re-evaluates inside it, and its
'     classifier then reads Date1904 on the workbook whose Date1904 is being
'     changed. That reentrant read crashed Excel rather than raising. With
'     manual calculation the toggle completes first and the only
'     recalculations are the explicit Application.Calculate calls.
'   - The scratch workbook is closed with SaveChanges:=False on every exit
'     path, including a raise.
'   - Formulas use ISO text, so the test is about caller classification and
'     not about serial interpretation.
'
' UPDATED
'   2026-09-01
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const ERR_NA        As Long = 2042  'Excel #N/A error number
    Const EXPECT_CHECKS As Long = 9     'Three cases, three assertions each
    Dim Scratch         As Workbook     'The scratch workbook, held by reference
    Dim Sheet           As Worksheet    'Its first sheet
    Dim Source          As String       'Source workbook name, quoted for formulas
    Dim PriorCalc       As XlCalculation 'Caller's calculation mode, restored on exit
    Dim CalcChanged     As Boolean      'TRUE once PriorCalc has been captured
    Dim Results         As Variant      'Report array
    Dim I               As Long         'Row cursor

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Fresh state for this run
        Set mFailures = New Collection
        mChecks = 0

    'Qualify UDF calls with this workbook, escaping any apostrophe in its name
        Source = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"

    'From here on, every exit must restore state and close the scratch workbook
        On Error GoTo Cleanup

    'Nothing may calculate except on explicit request. Captured before the
    'workbook exists so that Cleanup can always restore it.
        PriorCalc = Application.Calculation
        CalcChanged = True
        Application.Calculation = xlCalculationManual

    'Create and hold the scratch workbook by reference
        Set Scratch = Application.Workbooks.Add
        Set Sheet = Scratch.Worksheets(1)

'------------------------------------------------------------------------------
' 1900 WORKBOOK
'------------------------------------------------------------------------------
    'Start from a known state
        Scratch.Date1904 = False
        HostCase Sheet, Source, "1900", 1900, ERR_NA, False, True

'------------------------------------------------------------------------------
' 1904 WORKBOOK: SAME WORKBOOK, ORDINARY RECALCULATION
'------------------------------------------------------------------------------
    'Toggle the same workbook and recalculate normally, not a full rebuild
        Scratch.Date1904 = True
        HostCase Sheet, Source, "1904", 1904, ERR_NA, True, False

'------------------------------------------------------------------------------
' BACK TO 1900: THE VOLATILE DIAGNOSTIC MUST FOLLOW
'------------------------------------------------------------------------------
    'Toggle back and confirm the diagnostic tracks the change
        Scratch.Date1904 = False
        HostCase Sheet, Source, "1900 again", 1900, ERR_NA, False, False

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
Cleanup:
    'A raise reaching here is itself a failure to report
        If Err.Number <> 0 Then
            Record "host/runner", "unexpected runtime error " & CStr(Err.Number) & ": " & Err.Description
            Err.Clear
        End If

    'Close exactly the scratch workbook, never anything else, then restore the
    'caller's calculation mode. Order matters: closing under manual calculation
    'means no recalculation can run against a workbook being torn down.
        On Error Resume Next
        If Not Scratch Is Nothing Then Scratch.Close SaveChanges:=False
        If CalcChanged Then Application.Calculation = PriorCalc
        On Error GoTo 0

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'A count other than the fixed number this runner makes means its state was
    'overwritten mid-run, so nothing it reports can be trusted. Say so loudly.
        If mChecks <> EXPECT_CHECKS Then
            Record "host/runner state", "expected " & CStr(EXPECT_CHECKS) & _
                   " assertions but counted " & CStr(mChecks) & _
                   "; another run interrupted this one and its results were lost"
        End If

    'Summary first, then one line per failure
        Debug.Print "KPR host regression  checks: " & CStr(mChecks) & "  failures: " & CStr(mFailures.Count)
        For I = 1 To mFailures.Count
            Debug.Print "  FAIL  " & CStr(mFailures(I)(0)) & " : " & CStr(mFailures(I)(1))
        Next I

End Sub

Private Sub HostCase( _
    ByVal Sheet As Worksheet, _
    ByVal Source As String, _
    ByVal Label As String, _
    ByVal ExpectSystem As Long, _
    ByVal ErrNA As Long, _
    ByVal ExpectRefusal As Boolean, _
    ByVal WriteDiagnostic As Boolean)
'
' Writes the probe formulas into the scratch sheet, calculates, and asserts.
'
' The diagnostic in A1 is written ONCE, on the first case, and never touched
' again. Rewriting it would dirty the cell and mask the property under test:
' that a volatile zero-argument function refreshes on ordinary calculation
' after the date system changes. The value probes in A2 and A3 are rewritten
' on every case, because a non-volatile cell is not expected to notice a
' date-system toggle by itself; what they test is the guard's answer for a
' formula entered under each system.
'

'------------------------------------------------------------------------------
' WRITE PROBES
'------------------------------------------------------------------------------
    'The diagnostic, once only
        If WriteDiagnostic Then
            Sheet.Range("A1").Formula = "=" & Source & "KPR_Dates_HostDateSystem()"
        End If

    'A value function and a propagated #N/A, freshly entered under this system
        Sheet.Range("A2").Formula = "=" & Source & "KPR_Dates_BeginOfMonth(""2026-03-15"")"
        Sheet.Range("A3").Formula = "=" & Source & "KPR_Dates_BeginOfMonth(NA())"

    'Ordinary calculation only, requested explicitly under manual mode. A full
    'rebuild would mask a stale volatile cell.
        Application.Calculate

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'The diagnostic reports the current date system
        mChecks = mChecks + 1
        If VarType(Sheet.Range("A1").Value) = vbError Then
            Record "host/" & Label & " diagnostic", "expected " & CStr(ExpectSystem) & ", got an Excel error"
        ElseIf CLng(Sheet.Range("A1").Value) <> ExpectSystem Then
            Record "host/" & Label & " diagnostic", "expected " & CStr(ExpectSystem) & ", got " & CStr(Sheet.Range("A1").Value)
        End If

    'A value function proceeds under 1900 and is refused under 1904
        If ExpectRefusal Then
            AssertErrorValue "host/" & Label & " value function refused", Sheet.Range("A2").Value, ErrNA
        Else
            AssertDateResult "host/" & Label & " value function proceeds", Sheet.Range("A2").Value, DateSerial(2026, 3, 1)
        End If

    'A propagated #N/A is #N/A in both systems; it keeps its own label because
    'under 1904 the value is indistinguishable from host refusal
        AssertErrorValue "host/" & Label & " propagated NA", Sheet.Range("A3").Value, ErrNA

End Sub

'
'------------------------------------------------------------------------------
'
'                                  ASSERTIONS
'
'------------------------------------------------------------------------------
'

Private Sub AssertDateCondition( _
    ByVal Label As String, _
    ByVal ScalarIn As Variant, _
    ByVal ExpectOk As Boolean, _
    ByVal ExpectCondition As String)
'
' Asserts the exact condition identifier returned by the date parser.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parsed      As Date             'Parser output on success
    Dim Cond        As KPR_Condition    'Reported condition
    Dim Ok          As Boolean          'Parser return

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Count every assertion, passing or failing
        mChecks = mChecks + 1

    'Call the parser under test
        Ok = TryParseDateScalar(ScalarIn, Parsed, Cond)

    'Compare the outcome
        If Ok <> ExpectOk Then
            Record Label, "expected ok=" & CStr(ExpectOk) & " got " & CStr(Ok)
        ElseIf ConditionName(Cond) <> ExpectCondition Then
            Record Label, "expected " & ExpectCondition & " got " & ConditionName(Cond)
        End If

End Sub

Private Sub AssertDateValue( _
    ByVal Label As String, _
    ByVal ScalarIn As Variant, _
    ByVal ExpectDate As Date)
'
' Asserts that an accepted scalar normalizes to an exact date-only value.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parsed      As Date             'Parser output
    Dim Cond        As KPR_Condition    'Reported condition

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Count every assertion
        mChecks = mChecks + 1

    'Accept and compare
        If Not TryParseDateScalar(ScalarIn, Parsed, Cond) Then
            Record Label, "expected acceptance, got " & ConditionName(Cond)
        ElseIf Parsed <> ExpectDate Then
            Record Label, "expected " & Format$(ExpectDate, "yyyy-mm-dd") & _
                          " got " & Format$(Parsed, "yyyy-mm-dd")
        End If

End Sub

Private Sub AssertLongCondition( _
    ByVal Label As String, _
    ByVal ScalarIn As Variant, _
    ByVal ExpectOk As Boolean, _
    ByVal ExpectCondition As String)
'
' Asserts the exact condition identifier returned by the integer parser.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parsed      As Long             'Parser output on success
    Dim Cond        As KPR_Condition    'Reported condition
    Dim Ok          As Boolean          'Parser return

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Count every assertion
        mChecks = mChecks + 1

    'Call the parser under test
        Ok = TryParseLongScalar(ScalarIn, Parsed, Cond)

    'Compare the outcome
        If Ok <> ExpectOk Then
            Record Label, "expected ok=" & CStr(ExpectOk) & " got " & CStr(Ok)
        ElseIf ConditionName(Cond) <> ExpectCondition Then
            Record Label, "expected " & ExpectCondition & " got " & ConditionName(Cond)
        End If

End Sub

Private Sub AssertErrorValue( _
    ByVal Label As String, _
    ByVal Actual As Variant, _
    ByVal ExpectErrorNumber As Long)
'
' Asserts that a facade result is a specific native Excel error value.
'

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Count every assertion
        mChecks = mChecks + 1

    'The result must be an error at all
        If VarType(Actual) <> vbError Then
            Record Label, "expected an Excel error, got a value"

    'And it must be the expected one
        ElseIf CLng(Actual) <> ExpectErrorNumber Then
            Record Label, "expected error " & CStr(ExpectErrorNumber) & _
                          " got " & CStr(CLng(Actual))
        End If

End Sub

Private Sub AssertDateResult( _
    ByVal Label As String, _
    ByVal Actual As Variant, _
    ByVal ExpectDate As Date)
'
' Asserts that a facade result is an exact date value.
'

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Count every assertion
        mChecks = mChecks + 1

    'Compare only when the result is not an error
        If VarType(Actual) = vbError Then
            Record Label, "expected a date, got Excel error " & CStr(CLng(Actual))
        ElseIf CDate(Actual) <> ExpectDate Then
            Record Label, "expected " & Format$(ExpectDate, "yyyy-mm-dd") & _
                          " got " & Format$(CDate(Actual), "yyyy-mm-dd")
        End If

End Sub

Private Sub Record( _
    ByVal Label As String, _
    ByVal Detail As String)
'
' Records one failing case.
'

'------------------------------------------------------------------------------
' APPEND
'------------------------------------------------------------------------------
    'Store label and detail as a two-element array
        mFailures.Add Array(Label, Detail)

End Sub

Private Function ConditionName( _
    ByVal Condition As KPR_Condition) _
    As String
'
' Maps an internal enum member to the semantic identifier used by fixtures and
' evidence. Enum numbers never leave the implementation.
'

'------------------------------------------------------------------------------
' MAP
'------------------------------------------------------------------------------
    Select Case Condition
        Case KPR_COND_NONE:                     ConditionName = "NONE"
        Case KPR_COND_DATE_TEXT_FORMAT:         ConditionName = "DATE_TEXT_FORMAT"
        Case KPR_COND_DATE_TEXT_LOCALE:         ConditionName = "DATE_TEXT_LOCALE"
        Case KPR_COND_DATE_TEXT_NUMERIC:        ConditionName = "DATE_TEXT_NUMERIC"
        Case KPR_COND_DATE_TEXT_IMPOSSIBLE:     ConditionName = "DATE_TEXT_IMPOSSIBLE"
        Case KPR_COND_DATE_TYPE_REJECTED:       ConditionName = "DATE_TYPE_REJECTED"
        Case KPR_COND_DATE_WINDOW:              ConditionName = "DATE_WINDOW"
        Case KPR_COND_INPUT_BLANK_REQUIRED:     ConditionName = "INPUT_BLANK_REQUIRED"
        Case KPR_COND_INPUT_ERROR_PROPAGATED:   ConditionName = "INPUT_ERROR_PROPAGATED"
        Case KPR_COND_INTEGER_FRACTION:         ConditionName = "INTEGER_FRACTION"
        Case KPR_COND_INTEGER_TYPE_REJECTED:    ConditionName = "INTEGER_TYPE_REJECTED"
        Case KPR_COND_INTEGER_RANGE:            ConditionName = "INTEGER_RANGE"
        Case KPR_COND_DOMAIN_YEAR:              ConditionName = "DOMAIN_YEAR"
        Case KPR_COND_DOMAIN_MONTH:             ConditionName = "DOMAIN_MONTH"
        Case KPR_COND_DOMAIN_WEEKDAY:           ConditionName = "DOMAIN_WEEKDAY"
        Case KPR_COND_DOMAIN_OCCURRENCE:        ConditionName = "DOMAIN_OCCURRENCE"
        Case KPR_COND_OCCURRENCE_ABSENT:        ConditionName = "OCCURRENCE_ABSENT"
        Case KPR_COND_RESULT_WINDOW:            ConditionName = "RESULT_WINDOW"
        Case KPR_COND_CONTROL_TYPE_REJECTED:    ConditionName = "CONTROL_TYPE_REJECTED"
        Case KPR_COND_CONTROL_ERROR_PROPAGATED: ConditionName = "CONTROL_ERROR_PROPAGATED"
        Case KPR_COND_SHAPE_UNSUPPORTED:        ConditionName = "SHAPE_UNSUPPORTED"
        Case KPR_COND_HOST_DATE1904:            ConditionName = "HOST_DATE1904"
        Case KPR_COND_HOST_UNRESOLVED:          ConditionName = "HOST_UNRESOLVED"
        Case KPR_COND_CONTROL_TOKEN_UNKNOWN:    ConditionName = "CONTROL_TOKEN_UNKNOWN"
        Case KPR_COND_PILLAR_TYPE_REJECTED:     ConditionName = "PILLAR_TYPE_REJECTED"
        Case KPR_COND_PILLAR_TOKEN_MALFORMED:   ConditionName = "PILLAR_TOKEN_MALFORMED"
        Case KPR_COND_PILLAR_DUPLICATE_UNIT:    ConditionName = "PILLAR_DUPLICATE_UNIT"
        Case KPR_COND_PILLAR_ALIAS_SIGNED:      ConditionName = "PILLAR_ALIAS_SIGNED"
        Case KPR_COND_PILLAR_AGGREGATE_RANGE:   ConditionName = "PILLAR_AGGREGATE_RANGE"
        Case Else:                              ConditionName = "UNKNOWN"
    End Select

End Function

'
'------------------------------------------------------------------------------
'
'                            SUITES - DATE CLASSIFICATION
'
'------------------------------------------------------------------------------
'

Private Sub Run_DateTypeCases()
'
' Accepted and rejected scalar TYPES at a date position.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Obj         As Collection       'Representative rejected object type

'------------------------------------------------------------------------------
' ACCEPTED TYPES
'------------------------------------------------------------------------------
    'Native Date, with and without a time component
        AssertDateValue "date/native", DateSerial(2026, 3, 15), DateSerial(2026, 3, 15)
        AssertDateValue "date/native with time", DateSerial(2026, 3, 15) + 0.75, DateSerial(2026, 3, 15)

    'Every native numeric Variant subtype accepted by the parser
        AssertDateValue "serial/byte", CByte(255), CDate(255)
        AssertDateValue "serial/integer", CInt(1000), CDate(1000)
        AssertDateValue "serial/long", CLng(46096), DateSerial(2026, 3, 15)
        AssertDateValue "serial/single", CSng(46096.75), DateSerial(2026, 3, 15)
        AssertDateValue "serial/integral", 46096#, DateSerial(2026, 3, 15)
        AssertDateValue "serial/fractional", 46096.75, DateSerial(2026, 3, 15)
        AssertDateValue "serial/currency", CCur(46096.75), DateSerial(2026, 3, 15)
        AssertDateValue "serial/decimal", CDec(46096.75), DateSerial(2026, 3, 15)
#If VBA7 Then
        AssertDateValue "serial/longlong", CLngLng(46096), DateSerial(2026, 3, 15)
#End If

    'Exact ISO text
        AssertDateValue "iso/exact", "2026-03-15", DateSerial(2026, 3, 15)
        AssertDateValue "iso/leap day", "2024-02-29", DateSerial(2024, 2, 29)

'------------------------------------------------------------------------------
' REJECTED TYPES
'------------------------------------------------------------------------------
    'A blank required cell is not a zero serial
        AssertDateCondition "date/empty", Empty, False, "INPUT_BLANK_REQUIRED"

    'Boolean would otherwise coerce to serial -1
        AssertDateCondition "date/boolean true", True, False, "DATE_TYPE_REJECTED"
        AssertDateCondition "date/boolean false", False, False, "DATE_TYPE_REJECTED"

    'Null and error payloads never parse
        AssertDateCondition "date/null", Null, False, "DATE_TYPE_REJECTED"
        AssertDateCondition "date/error payload", CVErr(xlErrNA), False, "DATE_TYPE_REJECTED"

    'Objects are rejected rather than dereferenced through a default member
        Set Obj = New Collection
        AssertDateCondition "date/object", Obj, False, "DATE_TYPE_REJECTED"

End Sub

Private Sub Run_DateTextCases()
'
' Text is accepted in exactly one form. Every other shape is classified.
'

'------------------------------------------------------------------------------
' LOCALE AND NUMERIC TEXT
'------------------------------------------------------------------------------
    'Locale-formatted dates, in both day-first and month-first order
        AssertDateCondition "text/dmy slash", "31/12/2026", False, "DATE_TEXT_LOCALE"
        AssertDateCondition "text/mdy slash", "12/31/2026", False, "DATE_TEXT_LOCALE"
        AssertDateCondition "text/dotted", "31.12.2026", False, "DATE_TEXT_LOCALE"

    'Numeric-looking text is never reinterpreted as a serial
        AssertDateCondition "text/serial digits", "46096", False, "DATE_TEXT_NUMERIC"
        AssertDateCondition "text/serial decimal", "46096.75", False, "DATE_TEXT_NUMERIC"
        AssertDateCondition "text/small integer", "61", False, "DATE_TEXT_NUMERIC"

'------------------------------------------------------------------------------
' MALFORMED ISO
'------------------------------------------------------------------------------
    'Whitespace is a rejection, not something to clean up
        AssertDateCondition "text/leading space", " 2026-03-15", False, "DATE_TEXT_FORMAT"
        AssertDateCondition "text/trailing space", "2026-03-15 ", False, "DATE_TEXT_FORMAT"

    'Empty and short forms
        AssertDateCondition "text/empty", "", False, "DATE_TEXT_FORMAT"
        AssertDateCondition "text/omitted zeros", "2026-3-15", False, "DATE_TEXT_FORMAT"
        AssertDateCondition "text/no separators", "20260315", False, "DATE_TEXT_NUMERIC"

    'Time suffix and alternate separators
        AssertDateCondition "text/time suffix", "2026-03-15T00:00", False, "DATE_TEXT_FORMAT"
        AssertDateCondition "text/underscores", "2026_03_15", False, "DATE_TEXT_FORMAT"
        AssertDateCondition "text/non-digit", "20X6-03-15", False, "DATE_TEXT_FORMAT"

'------------------------------------------------------------------------------
' IMPOSSIBLE DATES
'------------------------------------------------------------------------------
    'Exact shape, impossible value: rollover must never pass silently
        AssertDateCondition "iso/feb 29 common year", "2025-02-29", False, "DATE_TEXT_IMPOSSIBLE"
        AssertDateCondition "iso/feb 30", "2024-02-30", False, "DATE_TEXT_IMPOSSIBLE"
        AssertDateCondition "iso/april 31", "2026-04-31", False, "DATE_TEXT_IMPOSSIBLE"
        AssertDateCondition "iso/month 13", "2026-13-01", False, "DATE_TEXT_IMPOSSIBLE"
        AssertDateCondition "iso/month 00", "2026-00-10", False, "DATE_TEXT_IMPOSSIBLE"
        AssertDateCondition "iso/day 00", "2026-03-00", False, "DATE_TEXT_IMPOSSIBLE"

    'Century leap rule: 1900 is not a leap year, 2000 is
        AssertDateCondition "iso/1900-02-29", "1900-02-29", False, "DATE_TEXT_IMPOSSIBLE"
        AssertDateCondition "text/dotted numeric shape", "31.12", False, "DATE_TEXT_NUMERIC"
        AssertDateCondition "text/two decimal points", "1.2.3", False, "DATE_TEXT_LOCALE"
        AssertDateCondition "text/sign only", "-", False, "DATE_TEXT_FORMAT"
        AssertDateCondition "text/signed serial", "-61", False, "DATE_TEXT_NUMERIC"
        AssertDateCondition "text/thousands separator", "46,096", False, "DATE_TEXT_FORMAT"
        AssertDateValue "iso/2000-02-29", "2000-02-29", DateSerial(2000, 2, 29)

End Sub

Private Sub Run_DateWindowCases()
'
' The supported window is closed at both ends and gated after time removal.
'

'------------------------------------------------------------------------------
' BOUNDARIES ACCEPTED
'------------------------------------------------------------------------------
    'Both bounds are inside the window
        AssertDateValue "window/min iso", "1900-03-01", DateSerial(1900, 3, 1)
        AssertDateValue "window/max iso", "9999-12-31", DateSerial(9999, 12, 31)
        AssertDateValue "window/min serial", 61#, DateSerial(1900, 3, 1)
        AssertDateValue "window/max serial", 2958465#, DateSerial(9999, 12, 31)

    'A time component on the upper bound still resolves to that day
        AssertDateValue "window/max serial with time", 2958465.999, DateSerial(9999, 12, 31)

'------------------------------------------------------------------------------
' ADJACENT VALUES REJECTED
'------------------------------------------------------------------------------
    'Serial 60 is the fictitious 1900-02-29 and is excluded by design
        AssertDateCondition "window/serial 60", 60#, False, "DATE_WINDOW"
        AssertDateCondition "window/serial 59", 59#, False, "DATE_WINDOW"
        AssertDateCondition "window/serial 0", 0#, False, "DATE_WINDOW"
        AssertDateCondition "window/negative serial", -1#, False, "DATE_WINDOW"

    'Just past the upper bound
        AssertDateCondition "window/serial 2958466", 2958466#, False, "DATE_WINDOW"

    'Below the floor year, expressed as exact ISO text
        AssertDateCondition "window/iso 1900-02-28", "1900-02-28", False, "DATE_WINDOW"
        AssertDateCondition "window/iso year 0050", "0050-01-01", False, "DATE_WINDOW"
        AssertDateCondition "window/iso year 1899", "1899-12-31", False, "DATE_WINDOW"

End Sub

'
'------------------------------------------------------------------------------
'
'                        SUITES - INTEGERS AND CONTROLS
'
'------------------------------------------------------------------------------
'

Private Sub Run_IntegerCases()
'
' Integer arguments reject fractions, Boolean and text, and range precedes
' integrality.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Obj         As Collection       'Representative rejected object type

'------------------------------------------------------------------------------
' ACCEPTED
'------------------------------------------------------------------------------
    'Every native numeric Variant subtype accepted by the parser
        AssertLongCondition "long/byte", CByte(12), True, "NONE"
        AssertLongCondition "long/integer", CInt(12), True, "NONE"
        AssertLongCondition "long/long", CLng(12), True, "NONE"
        AssertLongCondition "long/single", CSng(12), True, "NONE"
        AssertLongCondition "long/double", CDbl(12), True, "NONE"
        AssertLongCondition "long/currency", CCur(12), True, "NONE"
        AssertLongCondition "long/decimal", CDec(12), True, "NONE"
#If VBA7 Then
        AssertLongCondition "long/longlong", CLngLng(12), True, "NONE"
#End If

    'Integral values across the sign range
        AssertLongCondition "long/zero", 0, True, "NONE"
        AssertLongCondition "long/negative", -12, True, "NONE"
        AssertLongCondition "long/max", 2147483647#, True, "NONE"
        AssertLongCondition "long/min", -2147483648#, True, "NONE"

'------------------------------------------------------------------------------
' REJECTED
'------------------------------------------------------------------------------
    'No truncation and no rounding, including at the banker's-rounding tie
        AssertLongCondition "long/fraction up", 2.6, False, "INTEGER_FRACTION"
        AssertLongCondition "long/fraction tie", 2.5, False, "INTEGER_FRACTION"
        AssertLongCondition "long/fraction down", 2.4, False, "INTEGER_FRACTION"
        AssertLongCondition "long/tiny fraction", 3.0000001, False, "INTEGER_FRACTION"

    'Range precedes integrality when both apply
        AssertLongCondition "long/over range", 2147483648#, False, "INTEGER_RANGE"
        AssertLongCondition "long/under range", -2147483649#, False, "INTEGER_RANGE"
        AssertLongCondition "long/over range fractional", 2147483648.5, False, "INTEGER_RANGE"

    'Types that would otherwise coerce silently
        AssertLongCondition "long/boolean true", True, False, "INTEGER_TYPE_REJECTED"
        AssertLongCondition "long/boolean false", False, False, "INTEGER_TYPE_REJECTED"
        AssertLongCondition "long/numeric text", "12", False, "INTEGER_TYPE_REJECTED"
        AssertLongCondition "long/text", "twelve", False, "INTEGER_TYPE_REJECTED"
        AssertLongCondition "long/date", DateSerial(2026, 3, 15), False, "INTEGER_TYPE_REJECTED"
        AssertLongCondition "long/null", Null, False, "INTEGER_TYPE_REJECTED"
        AssertLongCondition "long/empty", Empty, False, "INPUT_BLANK_REQUIRED"
        Set Obj = New Collection
        AssertLongCondition "long/object", Obj, False, "INTEGER_TYPE_REJECTED"

End Sub

Private Sub Run_ControlCases()
'
' Optional controls accept a native Boolean or a default-selecting blank only.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parsed      As Boolean          'Control output
    Dim Cond        As KPR_Condition    'Reported condition
    Dim Obj         As Collection       'Representative rejected object type

'------------------------------------------------------------------------------
' ACCEPTED
'------------------------------------------------------------------------------
    'Blank selects the documented default, in both directions
        mChecks = mChecks + 1
        If Not TryParseBoolControl(Empty, True, Parsed, Cond) Then
            Record "control/empty defaults true", "expected acceptance"
        ElseIf Parsed <> True Then
            Record "control/empty defaults true", "expected the default True"
        End If

        mChecks = mChecks + 1
        If Not TryParseBoolControl(Empty, False, Parsed, Cond) Then
            Record "control/empty defaults false", "expected acceptance"
        ElseIf Parsed <> False Then
            Record "control/empty defaults false", "expected the default False"
        End If

    'An explicit native Boolean overrides the default
        mChecks = mChecks + 1
        If Not TryParseBoolControl(False, True, Parsed, Cond) Then
            Record "control/explicit false", "expected acceptance"
        ElseIf Parsed <> False Then
            Record "control/explicit false", "expected False to override the default"
        End If

'------------------------------------------------------------------------------
' REJECTED
'------------------------------------------------------------------------------
    'Numeric and textual stand-ins are rejected, not coerced
        AssertControlRejected "control/numeric one", 1
        AssertControlRejected "control/numeric zero", 0
        AssertControlRejected "control/text TRUE", "TRUE"
        AssertControlRejected "control/text true", "true"
        AssertControlRejected "control/null", Null
        AssertControlRejected "control/date", DateSerial(2026, 3, 15)
        Set Obj = New Collection
        AssertControlRejected "control/object", Obj

End Sub

Private Sub AssertControlRejected( _
    ByVal Label As String, _
    ByVal ControlIn As Variant)
'
' Asserts that a control value is rejected as CONTROL_TYPE_REJECTED.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parsed      As Boolean          'Control output
    Dim Cond        As KPR_Condition    'Reported condition

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Count every assertion
        mChecks = mChecks + 1

    'Reject with the expected condition
        If TryParseBoolControl(ControlIn, True, Parsed, Cond) Then
            Record Label, "expected rejection, got acceptance"
        ElseIf ConditionName(Cond) <> "CONTROL_TYPE_REJECTED" Then
            Record Label, "expected CONTROL_TYPE_REJECTED got " & ConditionName(Cond)
        End If

End Sub

'
'------------------------------------------------------------------------------
'
'                          SUITES - PUBLIC BOUNDARY
'
'------------------------------------------------------------------------------
'

Private Sub Run_BoundaryCases()
'
' Window mapping, propagation and composition, asserted through the facade
' because these decisions belong to the public boundary rather than the parser.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const ERR_VALUE As Long = 2015      'Excel #VALUE! error number
    Const ERR_NUM   As Long = 2036      'Excel #NUM! error number
    Const ERR_NA    As Long = 2042      'Excel #N/A error number

'------------------------------------------------------------------------------
' ERROR MAPPING
'------------------------------------------------------------------------------
    'A malformed value is #VALUE!
        AssertErrorValue "facade/locale text", KPR_Dates_DayOfWeek("31/12/2026"), ERR_VALUE

    'An out-of-window value is #NUM!, not #VALUE!
        AssertErrorValue "facade/window below", KPR_Dates_DayOfWeek("1900-02-28"), ERR_NUM
        AssertErrorValue "facade/window serial 60", KPR_Dates_DayOfWeek(60#), ERR_NUM

    'Both bounds are accepted
        AssertDateResult "facade/min accepted", KPR_Dates_BeginOfMonth("1900-03-01"), DateSerial(1900, 3, 1)
        AssertDateResult "facade/max accepted", KPR_Dates_EndOfMonth("9999-12-31"), DateSerial(9999, 12, 31)

'------------------------------------------------------------------------------
' PROPAGATION
'------------------------------------------------------------------------------
    'An incoming error is returned verbatim, not collapsed into #VALUE!
        AssertErrorValue "facade/propagate NA", KPR_Dates_DayOfWeek(CVErr(xlErrNA)), ERR_NA
        AssertErrorValue "facade/propagate NUM", KPR_Dates_DayOfWeek(CVErr(xlErrNum)), ERR_NUM
        AssertErrorValue "facade/propagate in integer slot", _
                         KPR_Dates_AddWeeks(DateSerial(2026, 3, 15), CVErr(xlErrNA)), ERR_NA
        AssertErrorValue "facade/propagate in control slot", _
                         KPR_Dates_DayOfWeek(DateSerial(2026, 3, 15), CVErr(xlErrNA)), ERR_NA

    'The first failing argument in signature order determines the result
        AssertErrorValue "facade/first argument wins", _
                         KPR_Dates_AddWeeks("31/12/2026", CVErr(xlErrNA)), ERR_VALUE

'------------------------------------------------------------------------------
' INTEGER AND CONTROL ARGUMENTS AT THE BOUNDARY
'------------------------------------------------------------------------------
    'A fractional shift is rejected rather than silently rounded
        AssertErrorValue "facade/fractional weeks", _
                         KPR_Dates_AddWeeks(DateSerial(2026, 3, 15), 2.5), ERR_VALUE

    'An out-of-Long shift is #NUM!
        AssertErrorValue "facade/out of long weeks", _
                         KPR_Dates_AddWeeks(DateSerial(2026, 3, 15), 2147483648#), ERR_NUM

    'A coercible stand-in for a Boolean control is rejected
        AssertErrorValue "facade/numeric control", _
                         KPR_Dates_DayOfWeek(DateSerial(2026, 3, 15), 1), ERR_VALUE

    'An omitted control still selects its default
        mChecks = mChecks + 1
        If KPR_Dates_DayOfWeek(DateSerial(2026, 3, 15)) <> _
           KPR_Dates_DayOfWeek(DateSerial(2026, 3, 15), True) Then
            Record "facade/omitted control default", "omitted and explicit True disagree"
        End If

'------------------------------------------------------------------------------
' SHAPE
'------------------------------------------------------------------------------
    'A non-scalar payload is rejected at the shape boundary
        AssertErrorValue "facade/array input", KPR_Dates_DayOfWeek(Array(1, 2)), ERR_VALUE

End Sub

Private Sub Run_MapperCases()
'
' ErrForCondition maps error conditions and refuses everything else.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const ERR_VALUE As Long = 2015      'Excel #VALUE! error number
    Const ERR_NUM   As Long = 2036      'Excel #NUM! error number

'------------------------------------------------------------------------------
' MAPPED CONDITIONS
'------------------------------------------------------------------------------
    'Contract-invalid input maps to #VALUE!
        AssertErrorValue "map/text format", ErrForCondition(KPR_COND_DATE_TEXT_FORMAT), ERR_VALUE
        AssertErrorValue "map/integer fraction", ErrForCondition(KPR_COND_INTEGER_FRACTION), ERR_VALUE
        AssertErrorValue "map/control type", ErrForCondition(KPR_COND_CONTROL_TYPE_REJECTED), ERR_VALUE

    'Out-of-domain maps to #NUM!
        AssertErrorValue "map/date window", ErrForCondition(KPR_COND_DATE_WINDOW), ERR_NUM
        AssertErrorValue "map/integer range", ErrForCondition(KPR_COND_INTEGER_RANGE), ERR_NUM
        AssertErrorValue "map/occurrence absent", ErrForCondition(KPR_COND_OCCURRENCE_ABSENT), ERR_NUM

'------------------------------------------------------------------------------
' UNMAPPED CONDITIONS RAISE
'------------------------------------------------------------------------------
    'The success sentinel has no error and must not answer #VALUE!
        AssertMapperRaises "map/none raises", KPR_COND_NONE

    'A propagation condition must not be mapped either: the incoming error is
    'the value to return, and discarding it is a defect
        AssertMapperRaises "map/propagated raises", KPR_COND_INPUT_ERROR_PROPAGATED

End Sub

Private Sub AssertMapperRaises( _
    ByVal Label As String, _
    ByVal Condition As KPR_Condition)
'
' Asserts that ErrForCondition raises rather than answering an error value.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ignored     As Variant      'Discarded result if no raise occurs
    Dim Raised      As Boolean      'Whether the expected raise happened

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Count every assertion
        mChecks = mChecks + 1

    'Contain the expected raise locally
        On Error Resume Next
        Ignored = ErrForCondition(Condition)
        Raised = (Err.Number <> 0)
        Err.Clear
        On Error GoTo 0

    'A silent mapping would hide an internal defect
        If Not Raised Then
            Record Label, "expected a raise, got a mapped error value"
        End If

End Sub
