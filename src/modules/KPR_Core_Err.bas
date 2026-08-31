Attribute VB_Name = "KPR_Core_Err"
'==============================================================================
' MODULE: KPR_Core_Err
'------------------------------------------------------------------------------
' PURPOSE
'   Internal boundary for native Excel error construction.
'
'   Every worksheet-facing failure value used by the KPR date layer is built
'   here, so that the set of errors the library can return is enumerable from
'   one place rather than recovered by grepping for CVErr.
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
'   Option Private Module. Members are Public only because VBA requires it for
'   cross-module calls inside the project. Nothing here is supported API, and
'   nothing here is visible to the worksheet, the Function Wizard, the macro
'   list or another VBA project.
'
' ALLOWED DEPENDENCIES
'   None. This module is the bottom of the dependency graph.
'
' NOTES
'   - ErrNA is provided but not yet called. Host-configuration #N/A arrives
'     with KPR_Dates_HostDateSystem in issue #17.
'   - Incoming-error propagation is deliberately absent. The migrated baseline
'     rejects an incoming Excel error rather than propagating it; issue #12
'     changes that, and this module gains the classification helpers then.
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
' RETURNS
'   Variant
'     CVErr(xlErrValue)
'
' ERROR POLICY
'   None. Construction only.
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
' RETURNS
'   Variant
'     CVErr(xlErrNum)
'
' ERROR POLICY
'   None. Construction only.
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
' RETURNS
'   Variant
'     CVErr(xlErrNA)
'
' ERROR POLICY
'   None. Construction only.
'
' NOTES
'   - Not called by the migrated surface. Reserved for the host date-system
'     refusal path introduced by issue #17.
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
