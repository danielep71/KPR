Attribute VB_Name = "KPR_Core_Array"
'==============================================================================
' MODULE: KPR_Core_Array
'------------------------------------------------------------------------------
' PURPOSE
'   Internal boundary for input shape: classifying an argument's shape and
'   unwrapping a single-valued wrapper down to the scalar it carries.
'
' WHY THIS EXISTS
'   The pre-split baseline mixed three concerns inside Parse_Date: unwrapping
'   shape, converting a scalar to a Date, and gating the supported window.
'   Shape lives here, value conversion lives in KPR_Core_Parse, and the window
'   belongs to the calendar domain in KPR_Core_Dates. Separating them is what
'   allows the array engine in issue #16 to traverse shapes without duplicating
'   value parsing.
'
' SCOPE (THIS REVISION)
'   - Array_Rank_1Or2   shape discriminator for 1-D and 2-D arrays
'   - TryUnwrapScalar   single-cell Range and 1x1 array unwrapping
'
'   Multi-element traversal, scalar expansion and exact-shape broadcasting are
'   NOT implemented here yet. They are owned by issue #16. This revision
'   preserves the baseline's scalar-only acceptance exactly.
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
'                              SHAPE CLASSIFICATION
'
'------------------------------------------------------------------------------
'

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

Public Function TryUnwrapScalar( _
    ByVal ValueIn As Variant, _
    ByRef ScalarOut As Variant) _
    As Boolean
'
'==============================================================================
'                                TryUnwrapScalar
'------------------------------------------------------------------------------
' PURPOSE
'   Reduces an incoming argument to the single scalar it carries, or fails.
'
' WHY THIS EXISTS
'   Worksheet arguments arrive as scalars, as single-cell Range objects, or as
'   1x1 array wrappers depending on how the formula was written. All three mean
'   one value. This routine is the single place that decides which wrappers are
'   equivalent to a scalar, so no caller re-implements shape logic.
'
' SIGNATURE
'   TryUnwrapScalar(ValueIn, ScalarOut) -> Boolean
'
' INPUTS
'   ValueIn
'     Any Variant argument, including Range objects and arrays.
'
' OUTPUTS
'   ScalarOut (ByRef)
'     Assigned ONLY on success, to the unwrapped scalar payload.
'
' RETURNS
'   Boolean
'     TRUE  => a single scalar was extracted; ScalarOut assigned
'     FALSE => the input is not single-valued; ScalarOut untouched
'
' BEHAVIOR
'   - An object is accepted only when it is a Range of exactly one cell, read
'     through Value2.
'   - An array is accepted only when it holds exactly one element, in either
'     one or two dimensions.
'   - Anything else, including a multi-cell Range and a non-Range object, fails.
'   - The scalar payload is NOT inspected here. Type acceptance is owned by
'     KPR_Core_Parse.
'
' ERROR POLICY
'   - Does not raise. All failure paths return FALSE.
'
' DEPENDENCIES
'   - Array_Rank_1Or2
'
' NOTES
'   - Rejecting a multi-element input is the baseline's scalar-only contract,
'     migrated unchanged. Issue #16 replaces rejection with traversal; this
'     routine keeps its meaning and becomes the 1x1 case of that engine.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V           As Variant      'Working value while wrappers are removed
    Dim Rng         As Range        'Bound Range when ValueIn is a Range object

    Dim ArrRank     As Long         'Array shape discriminator (1 or 2)
    Dim LB1         As Long         'Lower bound, dimension 1
    Dim UB1         As Long         'Upper bound, dimension 1
    Dim LB2         As Long         'Lower bound, dimension 2
    Dim UB2         As Long         'Upper bound, dimension 2

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime errors from property access and convert them to a FALSE return
        On Error GoTo Fail

    'Default return is failure unless we explicitly succeed
        TryUnwrapScalar = False

    'Work on a local copy to avoid repeated Variant indirection
        V = ValueIn

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
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once a single scalar has been isolated
        ScalarOut = V

    'Contract: TRUE only when ScalarOut was assigned
        TryUnwrapScalar = True
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving ScalarOut untouched
        TryUnwrapScalar = False

End Function
