# KPR v0.0.2 date-layer behavioural contract

Status: normative target contract for v0.0.2 implementation, registration,
regression tests, demo material, and user documentation.

This document freezes observable behaviour. The implementation plan defines
delivery order and repository architecture; this document defines what a
conforming date-layer call means. If descriptive material elsewhere conflicts
with this document, this document governs v0.0.2 behaviour. A contract change
requires an explicit, reviewed behavioural decision rather than an incidental
implementation change.

The pre-v0.0.2 `KPR_Dates_Days.bas` baseline does not yet conform to this
contract. In particular, this document does not assert that the current VBA
source imports, compiles, or passes regressions in Excel. Those claims require
exact-source Windows Excel evidence under issue #29.

## 1. Scope and terminology

The v0.0.2 date layer contains pure proleptic-Gregorian date primitives,
strict date and control parsing, date arithmetic, weekday locators, and pillar
conversion. Its supported date window is inclusive:

```text
1900-03-01 .. 9999-12-31
```

The lower bound deliberately excludes Excel serial 60, the fictitious
29-Feb-1900, and every earlier serial/date. All accepted date inputs and all
successful date outputs are normalized to date-only values.

In this document:

- **value argument** means a required argument whose values may determine the
  output shape. The first 21 functions below have one or more value arguments.
- **control argument** means an optional `Opt_` argument. A control never
  determines or expands the output shape.
- **scalar** means a non-array value or a one-cell/one-element wrapper.
- **non-scalar** means a multi-cell Range or multi-element in-memory array.
- **call-level error** means one scalar Excel error returned for the whole call;
  no output array is produced.
- **element-level error** means an Excel error stored only at the affected
  output position while valid neighbours continue to evaluate.

## 2. Exact supported public surface

Every public declaration returns `Variant` so it can return either its semantic
success type or a native Excel error. Every public argument is `ByVal`. Required
value arguments precede optional controls, and every optional control is last.

```vb
Public Function KPR_Dates_DayOfWeek(ByVal DateIn As Variant, Optional ByVal Opt_WeekBaseMonday As Variant = True) As Variant
Public Function KPR_Dates_DaysInMonth(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_DaysInYear(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_BeginOfMonth(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_EndOfMonth(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_BeginOfQuarter(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_EndOfQuarter(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_BeginOfYear(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_EndOfYear(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_IsMonthEnd(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_IsQuarterEnd(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_IsYearEnd(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_IsLeapYear(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_AddDays(ByVal DateIn As Variant, ByVal nDays As Variant) As Variant
Public Function KPR_Dates_AddWeeks(ByVal DateIn As Variant, ByVal nWeeks As Variant) As Variant
Public Function KPR_Dates_AddMonths(ByVal DateIn As Variant, ByVal nMonths As Variant, Optional ByVal Opt_KeepEOM As Variant = False) As Variant
Public Function KPR_Dates_AddYears(ByVal DateIn As Variant, ByVal nYears As Variant, Optional ByVal Opt_KeepEOM As Variant = False) As Variant
Public Function KPR_Dates_NthWeekdayOfMonth(ByVal YearIn As Variant, ByVal MonthIn As Variant, ByVal WdIndex As Variant, ByVal n As Variant, Optional ByVal Opt_WeekBaseMonday As Variant = True) As Variant
Public Function KPR_Dates_LastWeekdayOfMonth(ByVal YearIn As Variant, ByVal MonthIn As Variant, ByVal WdIndex As Variant, Optional ByVal Opt_WeekBaseMonday As Variant = True) As Variant
Public Function KPR_Dates_PillarFromDates(ByVal StartDate As Variant, ByVal EndDate As Variant, Optional ByVal Opt_Rounding As Variant = "NEAREST") As Variant
Public Function KPR_Dates_DateFromPillar(ByVal StartDate As Variant, ByVal Pillar As Variant) As Variant
Public Function KPR_Dates_HostDateSystem() As Variant
```

The plural pre-release name `KPR_Dates_DatesFromPillar` is replaced by singular
`KPR_Dates_DateFromPillar`. There is no compatibility alias. There are no
`_Spill` twins and no `KPR_Dates_Spill.bas` module.

| # | Supported name | Semantic scalar result | Vectorization |
| ---: | --- | --- | --- |
| 1 | `KPR_Dates_DayOfWeek` | `Long`, 1 through 7 | Array-capable |
| 2 | `KPR_Dates_DaysInMonth` | `Long`, 28 through 31 | Array-capable |
| 3 | `KPR_Dates_DaysInYear` | `Long`, 365 or 366 | Array-capable |
| 4 | `KPR_Dates_BeginOfMonth` | `Date` | Array-capable |
| 5 | `KPR_Dates_EndOfMonth` | `Date` | Array-capable |
| 6 | `KPR_Dates_BeginOfQuarter` | `Date` | Array-capable |
| 7 | `KPR_Dates_EndOfQuarter` | `Date` | Array-capable |
| 8 | `KPR_Dates_BeginOfYear` | `Date` | Array-capable |
| 9 | `KPR_Dates_EndOfYear` | `Date` | Array-capable |
| 10 | `KPR_Dates_IsMonthEnd` | `Boolean` | Array-capable |
| 11 | `KPR_Dates_IsQuarterEnd` | `Boolean` | Array-capable |
| 12 | `KPR_Dates_IsYearEnd` | `Boolean` | Array-capable |
| 13 | `KPR_Dates_IsLeapYear` | `Boolean` | Array-capable |
| 14 | `KPR_Dates_AddDays` | `Date` | Array-capable |
| 15 | `KPR_Dates_AddWeeks` | `Date` | Array-capable |
| 16 | `KPR_Dates_AddMonths` | `Date` | Array-capable |
| 17 | `KPR_Dates_AddYears` | `Date` | Array-capable |
| 18 | `KPR_Dates_NthWeekdayOfMonth` | `Date` | Array-capable |
| 19 | `KPR_Dates_LastWeekdayOfMonth` | `Date` | Array-capable |
| 20 | `KPR_Dates_PillarFromDates` | `String` | Array-capable |
| 21 | `KPR_Dates_DateFromPillar` | `Date` | Array-capable |
| 22 | `KPR_Dates_HostDateSystem` | `Long`, 1900 or 1904 | Scalar-only; no value argument |

Exactly these 22 names are supported public API in v0.0.2.

## 3. Date, integer, control, and token inputs

### 3.1 Date-value matrix

Date parsing is independent of Windows and Excel locale. Production code must
not use `CDate`, `IsDate`, or an equivalent permissive conversion to interpret
text.

| Input at one value position | Result |
| --- | --- |
| Incoming native Excel error (`CVErr`) | Propagate the same error value verbatim. |
| Native VBA `Date` | Accept, remove the time component, then apply the supported-window gate. |
| Native numeric scalar | Interpret under the call's documented 1900 serial contract, remove any fractional time component, then apply the supported-window gate. |
| Text exactly `YYYY-MM-DD` | Parse ASCII digits and hyphens only; validate year, month, day, month length, and leap day without rollover. |
| Locale-formatted text such as `31/12/2026` or `12/31/2026` | `#VALUE!`. |
| Numeric-looking text such as `61`, `45292`, or `45292.5` | `#VALUE!`; never reinterpret it as a serial. |
| Malformed ISO text, impossible ISO date, empty text, or whitespace-padded text | `#VALUE!`. |
| Required blank cell, `Empty`, `Null`, or Boolean | `#VALUE!`. |
| Range or supported in-memory array | Classify through the shape rules in section 5, then apply this matrix per element. |
| Non-Range object or unsupported array form | Call-level `#VALUE!`. |

ISO text is exactly ten characters. It accepts no alternate separator, omitted
zero, localized digit, time suffix, leading/trailing whitespace, or DateSerial
rollover. A syntactically exact ISO date outside the supported window returns
`#NUM!`; an impossible component such as `2025-02-29` returns `#VALUE!`.

### 3.2 Integer-value matrix

`nDays`, `nWeeks`, `nMonths`, `nYears`, `YearIn`, `MonthIn`, `WdIndex`, and `n`
use one strict integer parser at each output position.

| Input | Result |
| --- | --- |
| Native numeric value that is mathematically integral and inside the VBA `Long` range | Accept as `Long`. |
| Fractional numeric value | `#VALUE!`; no truncation, banker's rounding, or other rounding is permitted. |
| Numeric value outside the VBA `Long` range | `#NUM!`. |
| Boolean, Date, text (including numeric-looking text), blank, `Empty`, `Null`, or unsupported object | `#VALUE!`. |
| Incoming native Excel error | Propagate verbatim at that output position. |

After parsing, function-specific domains apply:

- `YearIn` is 1900 through 9999, but a requested result before 1900-03-01 is
  still outside the supported window and returns `#NUM!`.
- `MonthIn` must be 1 through 12, `WdIndex` must be 1 through 7, and occurrence
  `n` must be 1 through 5. A parsed integer outside one of these argument
  domains returns `#VALUE!`.
- Date-shift integers may be negative, zero, or positive. A valid shift that
  cannot produce a supported result returns `#NUM!`.

### 3.3 Optional-control matrix

An `Opt_` argument may be omitted, a scalar, or a 1x1 Range/array. It never
vectorizes. A multi-cell or multi-element optional argument returns call-level
`#VALUE!`.

| Control | Accepted values | Default | Invalid value |
| --- | --- | --- | --- |
| `Opt_WeekBaseMonday` | Native Boolean `True` or `False` only | `True` | Call-level `#VALUE!` |
| `Opt_KeepEOM` | Native Boolean `True` or `False` only | `False` | Call-level `#VALUE!` |
| `Opt_Rounding` | Text `NEAREST`, `FLOOR`, or `CEILING`, compared case-insensitively with no surrounding whitespace | `NEAREST` | Call-level `#VALUE!` |

An incoming scalar Excel error supplied as an optional control propagates
verbatim as a call-level result.

### 3.4 Pillar-token matrix

At each `Pillar` value position, `KPR_Dates_DateFromPillar` accepts case-
insensitive ASCII text with this grammar:

```text
[+|-] (ON | O/N | TN | T/N | ([0-9]+[YMWD])+)
```

The sign applies to the complete token. `ON`/`O/N` means `1D`; `TN`/`T/N`
means `2D`. Generic components may repeat and are accumulated. Year and month
components form one signed month delta; week and day components form one signed
day delta. The month delta is applied first with clip semantics, followed by
the exact day delta. No inner or surrounding whitespace, decimal quantity,
missing quantity, unknown unit, or partial parse is accepted. A malformed or
non-text token returns `#VALUE!`; a valid token whose aggregate or result is
outside the numerical/date domain returns `#NUM!`.

## 4. Caller and workbook date-system contract

Every public date-layer call classifies `Application.Caller` once at its public
boundary, before element conversion or traversal. Date-system authority comes
only from an identifiable worksheet caller:

1. If `Application.Caller` identifies a worksheet `Range`, read
   `Application.Caller.Parent.Parent.Date1904` from that exact workbook.
2. If that workbook uses the 1900 date system, evaluate normally.
3. If that workbook uses the 1904 date system, every value-taking public
   function returns one call-level library-produced `#N/A`. It must never
   return a plausible value shifted by 1,462 days.
4. If an identifiable worksheet host should be readable but its date system
   cannot be resolved reliably, return one call-level library-produced `#N/A`.
5. If no worksheet host can be identified, apply the documented 1900 serial
   contract. Never consult `ActiveWorkbook`, `ThisWorkbook`, or another
   unrelated workbook as a fallback.

Direct VBA, the Immediate window, `Application.Run`, and the regression harness
are the certified non-Range uses of the no-worksheet-host path. The caller owns
the interpretation of values it constructs. In particular, passing a Range as
a VBA argument does not turn a direct-VBA invocation into a worksheet-Range
caller.

A non-Range `Application.Caller` is not proof of direct VBA. Macro-attached
shapes may provide a String caller, and data validation, chart series, defined
names, and related Excel evaluation contexts may provide Error or other caller
forms. Those contexts are probed during Windows certification and are outside
the v0.0.2 compatibility claim.

### `KPR_Dates_HostDateSystem()`

`KPR_Dates_HostDateSystem()` is the sole deliberately volatile date-layer
function. It calls `Application.Volatile True` at entry and returns:

- `1900` for an identifiable 1900 worksheet caller;
- `1904` for an identifiable 1904 worksheet caller;
- `1900` when no worksheet host can be identified, under the documented 1900
  serial contract; or
- `#N/A` only when an identifiable worksheet host should be readable but its
  date system cannot be resolved reliably.

Changing the caller context or workbook date system and performing an ordinary
recalculation must re-evaluate the diagnostic; a full calculation rebuild is
not required. The other 21 functions remain non-volatile.

## 5. Scalar, array, shape, and capacity contract

### 5.1 Supported shapes

- A scalar or a 1x1 Range/array is scalar.
- A one-dimensional initialized VBA array is interpreted as a `1xN` row.
- A two-dimensional array or a contiguous single-area Range preserves its
  exact rows, columns, and orientation.
- A scalar value argument expands to the target shape selected by the
  non-scalar value arguments.
- Multiple non-scalar value arguments must have exactly identical row and
  column dimensions.
- There is no row/column outer product, implicit cross-broadcasting, jagged
  traversal, multi-area Range support, or higher-dimensional support.
- A supplied Range is evaluated at its requested dimensions. It is never
  intersected with or shortened to `UsedRange`.

An all-scalar call returns a scalar `Variant`, not a 1x1 array. A non-scalar
call returns a two-dimensional `Variant` array even when the resolved shape is
one row or one column.

### 5.2 Call-level and element-level outcomes

The public boundary applies these stages once per call:

1. classify the caller and enforce the host/date-system rule;
2. classify value/control shapes and reject unsupported wrappers;
3. validate scalar controls and resolve exact-shape broadcasting;
4. enforce the output-element cap; and
5. traverse and evaluate elements.

Consequently:

- a 1904 or unreadable identifiable worksheet host returns call-level `#N/A`;
- an unsupported input shape, shape mismatch, or non-scalar `Opt_` argument
  returns call-level `#VALUE!`;
- a resolved target of at most 100,000 elements is permitted;
- a resolved target of 100,001 or more elements returns call-level `#NUM!`;
- a blank or otherwise invalid required element returns `#VALUE!` only at that
  position;
- an element result outside the supported numerical/date domain returns
  `#NUM!` only at that position; and
- an incoming native error propagates unchanged at that position without
  suppressing valid neighbours.

Where more than one value argument at the same output position contains an
incoming error, the first error in signature order is propagated. Call-level
failures necessarily take precedence because no element traversal occurs.

## 6. Native-error taxonomy and provenance

Public functions return intentional native Excel errors, never message strings,
`MsgBox` output, or plausible-looking fallback values.

| Error | Library condition |
| --- | --- |
| `#VALUE!` | Uninterpretable or contract-invalid input: malformed/non-ISO text, impossible ISO date, prohibited type, blank required value, fractional integer, invalid Boolean/control/domain token, unsupported shape, or shape mismatch. |
| `#NUM!` | Well-formed value outside the supported numerical/date domain: date/result outside 1900-03-01 through 9999-12-31, integer outside `Long`, impossible valid shift, absent requested fifth weekday, oversized aggregate, or output target above 100,000 elements. |
| Library-produced `#N/A` | A result is unavailable in the identifiable worksheet host configuration, including a 1904 worksheet call or an unreadable identifiable worksheet date system. |
| Any incoming Excel error | Propagate the same native error value verbatim, subject to the call-level guards in section 5. |

A library-produced host-configuration `#N/A` and a propagated incoming `#N/A`
are value-identical Excel errors. Provenance cannot be recovered from the
returned value alone. Test evidence must retain distinct case IDs and
originating-condition metadata. For an identifiable worksheet caller, volatile
`KPR_Dates_HostDateSystem()` supplies context: `1904` identifies host refusal;
`1900` leaves propagation or another documented input condition as the source.

## 7. Function semantics

All functions use the parsing, host, array, and error rules above.

### 7.1 Boundaries and predicates

- `DayOfWeek` returns 1 through 7. With `Opt_WeekBaseMonday=True`, Monday is 1
  and Sunday is 7. With `False`, Sunday is 1 and Saturday is 7.
- `DaysInMonth` returns the Gregorian length of the containing month.
- `DaysInYear` returns 366 for a Gregorian leap year and 365 otherwise. A year
  is leap when divisible by 4 except centuries not divisible by 400.
- `BeginOfMonth` and `EndOfMonth` return the first and last dates of the
  containing month.
- `BeginOfQuarter` and `EndOfQuarter` use calendar quarters Jan-Mar, Apr-Jun,
  Jul-Sep, and Oct-Dec.
- `BeginOfYear` and `EndOfYear` return 1 January and 31 December of the
  containing year.
- `IsMonthEnd`, `IsQuarterEnd`, and `IsYearEnd` test those same boundaries.
- `IsLeapYear` applies the Gregorian rule to the input date's year.

### 7.2 Date arithmetic

- `AddDays` adds the parsed signed integer as exact calendar days.
- `AddWeeks` adds exactly seven times the parsed signed integer as calendar
  days.
- `AddMonths` shifts by calendar months. In normal clip mode, retain the input
  day when it exists in the target month and otherwise use the target month's
  last day. With `Opt_KeepEOM=True`, an input that is its month's last day maps
  to the target month's last day; a non-EOM input still uses clip mode.
- `AddYears` is the same operation as `AddMonths` with a 12-month multiple, so
  clipping, leap-day handling, and optional EOM preservation cannot diverge.

Every intermediate and final result is range-gated before conversion to a VBA
`Date`. Arithmetic overflow or a result outside the supported window returns
`#NUM!` rather than a runtime error or rollover.

### 7.3 Weekday locators

`NthWeekdayOfMonth` returns occurrence `n` of `WdIndex` in `YearIn`/`MonthIn`
under the selected weekday base. A valid request for an occurrence that does
not exist, such as a fifth weekday absent from that month, returns `#NUM!`.

`LastWeekdayOfMonth` returns the final occurrence of `WdIndex` in the requested
year/month. Its argument domains and weekday-base interpretation are identical.
Any otherwise valid locator whose result is before 1900-03-01 returns `#NUM!`.

### 7.4 Pillar formatting and rounding

`PillarFromDates` returns `0D` for equal dates. An absolute interval shorter
than seven days returns the exact signed day token. Longer intervals compare
whole-week anchors and calendar-month anchors generated from `StartDate` in the
direction of `EndDate`. Month anchors use the same clip semantics as
`DateFromPillar`; month counts of 12 or more format as years plus residual
months.

Rounding operates on the absolute magnitude and then restores a leading minus
sign for a negative interval:

- `FLOOR` chooses the greatest candidate magnitude whose anchor does not pass
  `EndDate` in the interval direction.
- `CEILING` chooses the smallest candidate magnitude whose anchor reaches or
  passes `EndDate` in the interval direction.
- `NEAREST` chooses the anchor with the smallest absolute calendar-day distance
  from `EndDate`.

Tie-breaking is deterministic. Equivalent week and month anchors use the month
representation. For `NEAREST`, an equal-distance month candidate is preferred
to a week candidate; equal-distance adjacent month candidates choose the
ceiling month anchor. A mathematically possible remaining same-family tie also
chooses the ceiling anchor. Exact anchors remain exact under every mode.

`DateFromPillar` applies the grammar and month-then-day order in section 3.4.
Pillar parsing and formatting are mutually consistent for every emitted token.
Because `PillarFromDates` may round, this is deliberately not a general
invariant:

```text
DateFromPillar(StartDate, PillarFromDates(StartDate, EndDate)) = EndDate
```

Tests and documentation must not imply that equality for rounded intervals.

## 8. Excel-version compatibility

Scalar calls and multi-cell calls have separate compatibility claims:

- scalar calls may be claimed only for Excel versions exercised successfully
  during exact-source Windows certification;
- multi-cell calls may be tested, supported, and claimed only on dynamic-array
  Excel; and
- v0.0.2 makes no Ctrl+Shift+Enter or other legacy multi-cell compatibility
  claim.

The implementation performs no Excel-version detection and does not create or
simulate Excel's placement-level `#SPILL!` error. Spill placement remains
Excel's responsibility.

## 9. Supported API boundary and future namespaces

The 22 names in section 2 are the complete supported v0.0.2 calculation API.
VBA may require other procedures to be technically `Public` for cross-module
calls, `Application.Run`, MacroOptions, RibbonX, CommandBars, tests, or demo
generation. Technical visibility does not make those procedures supported API.

Internal `KPR_Core_*` modules use `Option Private Module`; helpers use the
narrowest practical visibility. Registration, UI callbacks, test runners, and
demo builders are explicitly unsupported infrastructure and contain no date
algorithms.

`KPR_Cal_*` is reserved for v0.0.3. v0.0.2 creates no calendar placeholder and
adds no calendar option to a pure `KPR_Dates_*` function.

The following are explicitly outside v0.0.2:

- calendars and calendar composition;
- weekend masks and holiday sets;
- business-day arithmetic;
- following, preceding, modified, nearest, or other roll conventions;
- duplicate `_Spill` functions or `KPR_Dates_Spill.bas`; and
- generated `.xlsm`, `.xlam`, `.xlsx`, or other Office binaries in git.

## 10. Issue #9 acceptance traceability

| Acceptance criterion | Normative coverage |
| --- | --- |
| Exact signatures, semantic return types, and vectorization for all 22 names | Section 2 |
| Input-type and native-error matrices contain no locale-dependent parsing | Sections 3 and 6 |
| Supported-window and caller/date-system policy are unambiguous | Sections 1 and 4 |
| `#VALUE!`, `#NUM!`, host `#N/A`, and verbatim propagated errors have explicit conditions | Sections 3, 5, and 6 |
| Host and propagated `#N/A` are value-identical and caller context is documented | Sections 4 and 6 |
| `HostDateSystem()` is volatile and ordinary-recalculation behaviour is specified | Section 4 |
| Certified direct-VBA uses are distinct from unsupported non-Range Excel contexts | Section 4 |
| Scalar expansion, controls, orientation, 1-D arrays, blanks, capacity, and per-element errors are specified | Sections 3 and 5 |
| Scalar and dynamic-array claims are separate, with no CSE claim | Section 8 |
| Pillar tie-breaking, negative intervals, and non-invariant round trips are defined | Sections 3.4 and 7.4 |
| Supported-API boundary and reserved `KPR_Cal_*` namespace are recorded | Section 9 |
| Calendars, weekends, holidays, business days, and roll conventions are out of scope | Section 9 |
