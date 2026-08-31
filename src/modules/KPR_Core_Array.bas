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
'   - Array_Rank        reports true array rank, 0 through 60
'   - TryUnwrapScalar   single-cell Range and 1x1 array unwrapping
'
'   Multi-element traversal, scalar expansion and exact-shape broadcasting are
'   NOT implemented here yet. They are owned by issue #16. This revision
'   preserves the baseline's scalar-only acceptance exactly.
'
' SHAPE VOCABULARY
'   Rank is reported rather than capped, so a caller can reject a shape it does
'   not support at its own boundary. Folding rank 3 and above into rank 2 would
'   let such an array reach code that indexes it with two subscripts, raising
'   error 9 inside the caller instead of failing cleanly here.
'
' TYPE HANDOFF
'   A single-cell Range is read through Value2, so a date-formatted cell leaves
'   this module as vbDouble rather than the vbDate that a Value read produces.
'
'   Nothing downstream depends on which of the two arrives: KPR_Core_Parse
'   resolves both through the same arithmetic path and returns the same
'   date-only result, verified on the host. Value2 is used because it applies
'   no Date or Currency coercion of its own, not because the parse layer
'   requires it.
'
'   What does matter is that the argument is never Let-copied before its type
'   is tested. A Let assignment from a Variant holding an object invokes that
'   object's default member, which for a Range is Value. That silently turns a
'   wrapper into a value before any shape test runs, which makes the non-Range
'   rejection unreachable except through error 438 and materializes a
'   multi-cell Range in full before CountLarge can reject it.
'
' VISIBILITY
'   Option Private Module. Members are declared Public so other modules in the
'   project can call them; the module-level Private setting is what keeps them
'   out of the Excel function list and off the project's external surface.
'   Nothing here is supported API.
'
' ALLOWED DEPENDENCIES
'   KPR_Core_Err. Not referenced in this revision; both members convert every
'   failure to a return value rather than raising.
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

Public Function Array_Rank( _
    ByVal ArrayIn As Variant) _
    As Long
'
'==============================================================================
'                                  Array_Rank
'------------------------------------------------------------------------------
' PURPOSE
'   Reports how many dimensions an array payload actually exposes.
'
' WHY THIS EXISTS
'   Callers must reject any shape they do not support, and VBA offers no
'   non-raising way to ask. Keeping the probe here means no caller has to
'   hold its own On Error Resume Next block.
'
'   Reporting true rank rather than a capped classification lets a caller
'   reject rank 3 and above at its own boundary. Folding 3D into 2D would
'   let such an array through to be indexed with two subscripts, raising
'   error 9 deep inside the caller instead.
'
' SIGNATURE
'   Array_Rank(ArrayIn) -> Long
'
' INPUTS
'   ArrayIn
'     Variant that may or may not hold an array payload. Never written to.
'
' RETURNS
'   Long
'     0  => non-array, or dynamic array that has never been ReDim'd
'     1  => 1D array, including an empty one returned by Array()
'     2  => 2D array
'     3+ => rank as declared, up to the VBA ceiling of 60
'
' ERROR POLICY
'   - Does NOT raise.
'   - Clears Err as a side effect; callers must not hold pending error
'     state across this call.
'   - Restores normal error handling for the remainder of THIS procedure
'     only. VBA scopes error state per procedure, so the caller's handler
'     is unaffected either way.
'
' NOTES
'   - Any result >= 1 guarantees LBound / UBound are safe for dimensions
'     1 to that value. No caller-side guard is required.
'   - Rank 0 covers both "not an array" and "array with no allocated
'     dimensions". Callers needing to separate these should use IsArray,
'     which does not raise.
'   - A jagged array is rank 1. Its elements being arrays is not visible
'     here.
'   - The probe is bounded at the VBA ceiling of 60 dimensions so that a
'     payload whose probe never fails cannot spin.
'   - ProbeBound is assigned and never read. It exists because LBound must
'     have a destination; only the success or failure of the call matters.
'     Removing it breaks the probe.
'
' UPDATED
'   2026-08-31
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ProbeDim    As Long      'Dimension currently under probe
    Dim ProbeBound  As Long      'Probe target; only success / failure matters
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default is 0 unless at least one dimension probes clean
        Array_Rank = 0
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Non-array input has no dimensions to count
        If Not IsArray(ArrayIn) Then Exit Function
'------------------------------------------------------------------------------
' DIMENSION PROBE
'------------------------------------------------------------------------------
    'Probe upward until a dimension is unavailable; VBA caps rank at 60
        On Error Resume Next
        For ProbeDim = 1 To 60
            Err.Clear
            ProbeBound = LBound(ArrayIn, ProbeDim)
            If Err.Number <> 0 Then Exit For
            Array_Rank = ProbeDim
        Next ProbeDim
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
'   - An array is accepted only when it is rank 1 or rank 2 and holds exactly
'     one element. Rank 0 and rank 3 or higher are rejected outright.
'   - A single element that is itself an object is rejected; it is a wrapper,
'     not a scalar payload.
'   - Anything else, including a multi-cell Range and a non-Range object, fails.
'   - The scalar payload is NOT inspected here. Type acceptance is owned by
'     KPR_Core_Parse.
'
' ERROR POLICY
'   - Does not raise. All failure paths return FALSE.
'   - Clears Err before returning FALSE, so no stale error state survives.
'
' DEPENDENCIES
'   - Array_Rank
'
' NOTES
'   - ValueIn is tested with IsObject directly and never Let-copied first. See
'     TYPE HANDOFF in the module header for what a Let copy would do.
'   - Rank is tested before any bound is read, so an unsupported shape is
'     rejected at the boundary rather than through the error handler.
'   - The wrapped-object rejection is a guard rather than a migrated rule. The
'     baseline had no equivalent, because its own Let copy meant an object
'     element was dereferenced before anything could test it.
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
    Dim Rng         As Range        'Bound Range when ValueIn is a Range object
    Dim Payload     As Variant      'Isolated scalar, assigned out only on success
    Dim ArrRank     As Long         'Array rank as reported by Array_Rank

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

'------------------------------------------------------------------------------
' ISOLATE PAYLOAD
'------------------------------------------------------------------------------
    'Test the argument itself; a Let copy would dereference the default member
        If IsObject(ValueIn) Then
            'Reject non-Range objects, Nothing included
                If Not TypeOf ValueIn Is Range Then GoTo Fail
            'Bind strongly typed
                Set Rng = ValueIn
            'Require exactly one cell before reading any content
                If Rng.CountLarge <> 1 Then GoTo Fail
            'Read through Value2 so no Date or Currency coercion is applied
                Payload = Rng.Value2

    'If an array arrived, accept only a 1x1 wrapper (scalar contract)
        ElseIf IsArray(ValueIn) Then
            'Classify the array shape once
                ArrRank = Array_Rank(ValueIn)
            'Reject unsupported shapes before any bound is read
                If ArrRank < 1 Or ArrRank > 2 Then GoTo Fail
            'Two-dimensional wrapper
                If ArrRank = 2 Then
                    'Capture bounds
                        LB1 = LBound(ValueIn, 1): UB1 = UBound(ValueIn, 1)
                        LB2 = LBound(ValueIn, 2): UB2 = UBound(ValueIn, 2)
                    'Reject anything larger than 1x1
                        If (UB1 <> LB1) Or (UB2 <> LB2) Then GoTo Fail
                    'A wrapped object is not a scalar payload
                        If IsObject(ValueIn(LB1, LB2)) Then GoTo Fail
                    'Unwrap the scalar payload
                        Payload = ValueIn(LB1, LB2)
            'One-dimensional wrapper
                Else
                    'Capture bounds
                        LB1 = LBound(ValueIn): UB1 = UBound(ValueIn)
                    'Reject vectors with more than one element
                        If UB1 <> LB1 Then GoTo Fail
                    'A wrapped object is not a scalar payload
                        If IsObject(ValueIn(LB1)) Then GoTo Fail
                    'Unwrap the scalar payload
                        Payload = ValueIn(LB1)
                End If

    'Neither object nor array: the argument is already a scalar
        Else
            Payload = ValueIn
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once a single scalar has been isolated
        ScalarOut = Payload
    'Contract: TRUE only when ScalarOut was assigned
        TryUnwrapScalar = True
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving ScalarOut untouched
        TryUnwrapScalar = False
    'Do not leave stale error state for the caller to observe
        Err.Clear

End Function
