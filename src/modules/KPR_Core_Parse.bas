Attribute VB_Name = "KPR_Core_Parse"
'==============================================================================
' MODULE: KPR_Core_Parse
'------------------------------------------------------------------------------
' PURPOSE
'   Internal boundary for value classification: converting one already-unwrapped
'   scalar into a normalized VBA Date, or failing.
'
' WHY THIS EXISTS
'   Worksheet arguments carry many runtime types for the same intent. This
'   module is the single place where the acceptance policy for a date value
'   lives, so callers never re-implement fragile type logic and no unhandled
'   runtime error can reach the worksheet.
'
' SCOPE (THIS REVISION)
'   - TryParseDateScalar   scalar -> date-only VBA Date
'
'   This revision migrates the baseline acceptance policy unchanged, including
'   two behaviours the frozen contract removes:
'
'     * date-like text is accepted under host locale rules (CDate / IsDate)
'     * numeric-looking text is reinterpreted as an Excel serial
'
'   Both are contract-invalid and are replaced by strict ISO-8601-only parsing
'   in issue #12, together with verbatim propagation of an incoming Excel error
'   in place of the outright rejection below. They are preserved here only so
'   that this migration changes structure without changing behaviour.
'
' SCOPE BOUNDARY
'   - Shape is NOT handled here. Callers pass a scalar already unwrapped by
'     KPR_Core_Array.TryUnwrapScalar.
'   - The supported date window is NOT applied here. The window is a calendar
'     fact owned by KPR_Core_Dates and is applied by the facade immediately
'     after parsing. The baseline applied it inside its combined parser; the
'     single-gate invariant is preserved, only its location changed.
'   - Integer and optional-control parsing are owned by issue #12 and are
'     deliberately absent rather than stubbed.
'
' VISIBILITY
'   Option Private Module. Members are Public only because VBA requires it for
'   cross-module calls inside the project. Nothing here is supported API.
'
' ALLOWED DEPENDENCIES
'   KPR_Core_Err.
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

'
'------------------------------------------------------------------------------
'
'                              DATE VALUE PARSING
'
'------------------------------------------------------------------------------
'

Public Function TryParseDateScalar( _
    ByVal ScalarIn As Variant, _
    ByRef ParsedDate As Date) _
    As Boolean
'
'==============================================================================
'                              TryParseDateScalar
'------------------------------------------------------------------------------
' PURPOSE
'   Converts a single already-unwrapped scalar into a normalized VBA Date
'   (date-only), returning a Boolean success flag.
'
' SIGNATURE
'   TryParseDateScalar(ScalarIn, ParsedDate) -> Boolean
'
' INPUTS
'   ScalarIn
'     Scalar Variant. Accepted in this revision:
'       - vbDate values
'       - numeric Excel serials
'       - date-like strings (host locale rules)      [contract-invalid, #12]
'       - numeric-looking strings treated as serials [contract-invalid, #12]
'
'     Rejected: Excel error values, Empty, Null, Boolean, and any scalar that
'     is neither a Date, a number, nor an acceptable string.
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
'   - The time component is removed from every accepted value, so a cell
'     holding NOW() resolves to its date part.
'   - No range gate is applied. The caller applies the supported window.
'
' ERROR POLICY
'   - Does not propagate VBA runtime errors to callers.
'   - Signals failure only through the Boolean return.
'
' NOTES
'   - Boolean is rejected deliberately: TRUE coerces to serial -1 in VBA, which
'     is neither a date the user meant nor a value inside the supported range.
'   - Rejecting an incoming Excel error is baseline behaviour, not contract
'     behaviour. The contract requires verbatim propagation; issue #12 moves
'     that decision to the caller, where an error value can be returned rather
'     than converted into a parse failure.
'   - Fix() truncates toward zero rather than flooring. Within the supported
'     window every serial is positive, so the two agree; negative serials fail
'     the caller's window gate regardless.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim VT          As VbVarType    'Cached VarType of the incoming scalar
    Dim Candidate   As Date         'Parsed candidate before assignment
    Dim X           As Double       'Numeric working value for serial inputs
    Dim S           As String       'Trimmed string token

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime coercion errors and convert them to a FALSE return
        On Error GoTo Fail

    'Default return is failure unless we explicitly succeed
        TryParseDateScalar = False

'------------------------------------------------------------------------------
' CLASSIFY SCALAR
'------------------------------------------------------------------------------
    'Classify the already-unwrapped payload
        VT = VarType(ScalarIn)

    Select Case VT

        Case vbError, vbEmpty, vbNull, vbBoolean
            'Excel errors, blanks and Boolean are never dates
                GoTo Fail

        Case vbDate
            'Native Date => strip any time component
                Candidate = DateValue(CDate(ScalarIn))

        Case vbString
            'Trim once for deterministic checks
                S = Trim$(CStr(ScalarIn))

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
                If Not IsNumeric(ScalarIn) Then GoTo Fail

            'Coerce once and strip the fractional day
                X = CDbl(ScalarIn)
                Candidate = CDate(Fix(X))

    End Select

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value has been classified
        ParsedDate = Candidate

    'Contract: TRUE only when ParsedDate was assigned
        TryParseDateScalar = True
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving ParsedDate untouched
        TryParseDateScalar = False

End Function
