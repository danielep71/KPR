Attribute VB_Name = "KPR_Core_Err"
'==============================================================================
' MODULE: KPR_Core_Err
'------------------------------------------------------------------------------
' PURPOSE
'   Internal boundary for native Excel error construction.
'
'   Every worksheet-facing failure value the KPR date layer CONSTRUCTS is built
'   here, so that the set of errors the library can originate is enumerable
'   from one place rather than recovered by grepping for CVErr.
'
' WHY THIS EXISTS
'   The pre-split baseline constructed CVErr(xlErrValue) and CVErr(xlErrNum)
'   inline at forty call sites. That made the error surface impossible to audit
'   and impossible to change coherently. Centralizing construction is the first
'   step; the behavioural error policy itself is owned by issue #13.
'
' SCOPE
'   - ErrValue   #VALUE!
'   - ErrNum     #NUM!
'   - ErrNA      #N/A
'
' VISIBILITY
'   Option Private Module. Members are declared Public so other modules in the
'   project can call them; the module-level Private setting is what keeps them
'   out of the Excel function list and off the project's external surface.
'   Nothing here is supported API, and nothing here is visible to the
'   worksheet, the Function Wizard, the macro list or another VBA project.
'
' ALLOWED DEPENDENCIES
'   No KPR module. This is the bottom of the project dependency graph.
'
'   It does depend on the host: CVErr is a VBA conversion function and the
'   xlErr* names are members of XlCVError in the Excel object library. Without
'   that reference the module does not compile. That is the only external
'   surface it touches.
'
' NOTES
'   - ErrNA is provided but not yet called. Host-configuration #N/A arrives
'     with KPR_Dates_HostDateSystem in issue #17.
'   - Incoming-error propagation is deliberately absent. The migrated baseline
'     rejects an incoming Excel error rather than propagating it; issue #12
'     changes that, and this module gains the classification helpers then.
'     Once propagation exists, a caller's own error will reach the sheet
'     without passing through this module, which is why the purpose above is
'     scoped to values the library constructs.
'   - Each member is a function rather than a constant because VBA cannot hold
'     an error value in a Const. The cost is one call per constructed error.
'   - A caller that assigns a failure value before knowing whether it will fail
'     pays that call on the success path too. On a per-cell surface evaluated
'     across a large range this is not free; construct at the failure point
'     unless a handler needs to substitute a different value first.
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
'                           NATIVE ERROR CONSTRUCTORS
'
'------------------------------------------------------------------------------
'

Public Function ErrValue() As Variant
'
'==============================================================================
'                                   ErrValue
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the native Excel #VALUE! error value.
'
' SIGNATURE
'   ErrValue() -> Variant
'
' RETURNS
'   Variant
'     CVErr(xlErrValue), a Variant of subtype vbError.
'
' ERROR POLICY
'   None. Construction only.
'
' NOTES
'   - The result must be Let-assigned, never Set. Testing it requires IsError;
'     comparing it to another value with = raises.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The input cannot be interpreted under the module contract
        ErrValue = CVErr(xlErrValue)

End Function


Public Function ErrNum() As Variant
'
'==============================================================================
'                                    ErrNum
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the native Excel #NUM! error value.
'
' SIGNATURE
'   ErrNum() -> Variant
'
' RETURNS
'   Variant
'     CVErr(xlErrNum), a Variant of subtype vbError.
'
' ERROR POLICY
'   None. Construction only.
'
' NOTES
'   - The contract classifies an out-of-window date as DATE_WINDOW and requires
'     this value. The migrated surface still returns #VALUE! there; issues #12
'     and #13 move it.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The input is well formed but the answer does not exist or is out of range
        ErrNum = CVErr(xlErrNum)

End Function


Public Function ErrNA() As Variant
'
'==============================================================================
'                                     ErrNA
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the native Excel #N/A error value.
'
' SIGNATURE
'   ErrNA() -> Variant
'
' RETURNS
'   Variant
'     CVErr(xlErrNA), a Variant of subtype vbError.
'
' ERROR POLICY
'   None. Construction only.
'
' NOTES
'   - Not called by the migrated surface. Reserved for the host date-system
'     refusal path introduced by issue #17.
'   - #N/A here means the library declines to answer, not that a lookup failed.
'     The distinction matters because a sheet cannot tell the two apart, so
'     the provenance of this value belongs in the contract rather than in a
'     cell comment.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The library refuses to answer for this caller or host configuration
        ErrNA = CVErr(xlErrNA)

End Function
