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
' TYPE HANDOFF
'   A single-cell Range arrives from KPR_Core_Array read through Value2, so a
'   date-formatted cell presents as vbDouble rather than vbDate. Both types
'   resolve through the same arithmetic path below and produce the same
'   date-only result, so the surface does not depend on which one arrives.
'
' VISIBILITY
'   Option Private Module. Members are declared Public so other modules in the
'   project can call them; the module-level Private setting is what keeps them
'   out of the Excel function list and off the project's external surface.
'   Nothing here is supported API.
'
' ALLOWED DEPENDENCIES
'   KPR_Core_Err. Not referenced in this revision; every failure is reported
'   through the Boolean return rather than as a worksheet error value.
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
'     Rejected: Excel error values, Empty, Null, Boolean, objects, and any
'     scalar that is neither a Date, a number, nor an acceptable string.
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
'   - Clears Err before returning FALSE, so no stale error state survives.
'
' NOTES
'   - Boolean is rejected deliberately: TRUE coerces to serial -1 in VBA, which
'     is neither a date the user meant nor a value inside the supported range.
'   - An object is rejected explicitly rather than left to fall through. Any
'     numeric test applied to an object reference invokes its default member,
'     which silently converts a wrapper into a value. Shape has already been
'     resolved upstream, so an object arriving here is a caller defect.
'   - Rejecting an incoming Excel error is baseline behaviour, not contract
'     behaviour. The contract requires verbatim propagation; issue #12 moves
'     that decision to the caller, where an error value can be returned rather
'     than converted into a parse failure.
'   - Fix is the correct operator for removing a time component, not merely
'     an equivalent one. VBA encodes a date serial as sign and magnitude: the
'     fraction is always time regardless of the sign of the whole part. Fix
'     truncates toward zero and so keeps the day; Int floors and would move a
'     negative serial back one day. Replacing Fix with Int is a defect even
'     though the two agree everywhere inside the supported window.
'   - The vbDate branch converts arithmetically rather than through DateValue.
'     DateValue takes a date expression that VBA coerces to String, so passing
'     an already-typed Date round-trips it through the host locale format for
'     no benefit. Working on the serial keeps the branch locale-independent
'     and identical to the numeric branch.
'   - IsNumeric is materially wider than "numeric-looking". It accepts hex
'     literals such as &HFF, exponent forms such as 1e5 and 1d5, currency
'     symbols and parenthesised negatives, so tokens no user would consider a
'     date currently parse as serials. This is migrated baseline behaviour and
'     is removed by the ISO-8601 policy in issue #12.
'   - Excel serials below 61 do not correspond to VBA serials, because the
'     Excel 1900 system counts a fictitious 29-Feb-1900. Nothing is done about
'     that here; every such value fails the caller's window gate, whose lower
'     bound is exactly serial 61.
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

        Case vbError, vbEmpty, vbNull, vbBoolean, vbObject
            'Excel errors, blanks, Boolean and object references are never dates
                GoTo Fail

        Case vbDate
            'Native Date => strip the time component on the serial
                Candidate = CDate(Fix(CDbl(ScalarIn)))

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

    'Do not leave stale error state for the caller to observe
        Err.Clear

End Function
