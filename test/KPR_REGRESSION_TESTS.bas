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
'       KPR_Tests_Run                 every suite; appears in the Alt+F8 macro
'                                     list and prints its report to the
'                                     Immediate window
'       KPR_Tests_RunSuite "integer"  one suite, printed the same way
'       KPR_Tests_RunAll()            returns a two-column array for a worksheet
'                                     or for programmatic use
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
'   KPR_Core_Parse and KPR_Core_Err, called directly to assert exact condition
'   classification, and KPR_DATES_DAYS for boundary behaviour. No other core is
'   reachable from here.
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

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain any unexpected error as a reported failure
        On Error GoTo Err_Handler

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

    'Return the array
        KPR_Tests_RunAll = Results
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'A raise reaching here is a defect, so report it rather than swallowing it
        KPR_Tests_RunAll = "unexpected runtime error " & CStr(Err.Number) & ": " & Err.Description

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

        Case "date-type":       Run_DateTypeCases
        Case "date-text":       Run_DateTextCases
        Case "date-window":     Run_DateWindowCases
        Case "integer":         Run_IntegerCases
        Case "control":         Run_ControlCases
        Case "boundary":        Run_BoundaryCases
        Case "mapper":          Run_MapperCases

        Case Else
            'Not in the registry
                RunSuite = False

    End Select

End Function

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
' ACCEPTED TYPES
'------------------------------------------------------------------------------
    'Native Date, with and without a time component
        AssertDateValue "date/native", DateSerial(2026, 3, 15), DateSerial(2026, 3, 15)
        AssertDateValue "date/native with time", DateSerial(2026, 3, 15) + 0.75, DateSerial(2026, 3, 15)

    'Numeric serial, integral and fractional; both normalize to the same day
        AssertDateValue "serial/integral", 46096#, DateSerial(2026, 3, 15)
        AssertDateValue "serial/fractional", 46096.75, DateSerial(2026, 3, 15)

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
' ACCEPTED
'------------------------------------------------------------------------------
    'Integral numerics across the sign range
        AssertLongCondition "long/zero", 0, True, "NONE"
        AssertLongCondition "long/positive", 12, True, "NONE"
        AssertLongCondition "long/negative", -12, True, "NONE"
        AssertLongCondition "long/integral double", 12#, True, "NONE"
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
