# v0.0.2 implementation plan

Status: approved for issue-driven implementation. No production implementation or release certification has started.

- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)
- Planning baseline: `main` at `54d64ab3624aa051b19f9677aa49a3782cd26c60`
- Last plan revision before this reconciliation: `6de661e8bb3a7561480fcaaf097b8af605071f57`
- Published protected baseline: `v0.0.1` at `abf38786eb48b3db1edced8ae26c756d9c7f5328`
- Version before candidate assembly: `0.0.1`

## Goal

Deliver the first functional KPR date layer as a source-first pre-release: pure Gregorian date primitives, strict parsing, one scalar/array-capable public surface, registration and Excel UI infrastructure, deterministic tests, a reproducible demo builder, and exact-source Windows Excel certification.

The milestone is not complete merely because hosted static checks pass. It may be tagged and published only after the exact candidate source has been imported, compiled and executed successfully in Windows Excel and the evidence has been attached.

## Milestone sequence and boundaries

| Milestone | Scope |
| --- | --- |
| `v0.0.2` | Date primitives, parsing, one array-capable API, MacroOptions, Ribbon, CommandBars, regression harness, demo and Windows certification |
| `v0.0.3` | Calendars: weekend masks, holiday sets and calendar composition |
| `v0.0.4` | Business-day arithmetic and roll conventions built on the calendar layer |

The `KPR_Cal_*` namespace is reserved now. v0.0.2 creates no calendar placeholder modules, includes no weekend masks or holiday arguments, and performs no business-day calculations.

## Frozen design decisions

| Area | v0.0.2 decision |
| --- | --- |
| Public surface | Exactly 22 supported `KPR_Dates_*` names in one facade. Every function with value arguments is array-capable. There are no `_Spill` twins. |
| Caller/date system | A worksheet `Range` caller is checked through its own workbook. A 1904 workbook returns call-level `#N/A`. A non-`Range` caller is treated as a direct VBA/infrastructure call and uses the documented 1900 serial contract. No unrelated workbook fallback is permitted. |
| Supported window | `1900-03-01 .. 9999-12-31`. Excel serial 60 and all earlier values are excluded. |
| String dates | ISO `YYYY-MM-DD` only. Locale-formatted and numeric-looking strings are rejected. |
| Errors | `#VALUE!` means uninterpretable or contract-invalid; `#NUM!` means well-formed but outside the supported domain; `#N/A` means unavailable in the host configuration. Incoming Excel errors propagate verbatim. |
| Arrays | Scalar expansion plus exact-shape broadcasting; shape and orientation preserved; one-dimensional VBA arrays are `1×N`; no outer broadcasting. |
| Optional arguments | All required value arguments first. Every `Opt_` argument is last and must be omitted, scalar or `1×1`; a larger optional argument is call-level `#VALUE!`. |
| Size guard | Maximum 100,000 output elements per call. A larger target shape returns call-level `#NUM!`. A supplied Range is never shortened through `UsedRange`. |
| Excel versions | Scalar calls are supported on the Excel versions actually certified for scalar use. Multi-cell calls are supported and claimed only on dynamic-array Excel. No CSE claim is made. |
| Pillar rounding | Case-insensitive `NEAREST` (default), `FLOOR` and `CEILING`. |
| Namespace | `KPR_Cal_*` is reserved for calendars. No calendar options are appended to pure date functions. |

## Supported API target

The following signatures are the target contract. All worksheet-reachable inputs and optional controls are `Variant` so native errors, blanks, ranges and arrays can be classified deliberately. Every public result is `Variant` so a native Excel error can be returned.

| # | Target signature | Successful scalar result |
| ---: | --- | --- |
| 1 | `KPR_Dates_DayOfWeek(DateIn As Variant, Optional Opt_WeekBaseMonday As Variant = True) As Variant` | `Long` 1–7 |
| 2 | `KPR_Dates_DaysInMonth(DateIn As Variant) As Variant` | `Long` 28–31 |
| 3 | `KPR_Dates_DaysInYear(DateIn As Variant) As Variant` | `Long` 365 or 366 |
| 4 | `KPR_Dates_BeginOfMonth(DateIn As Variant) As Variant` | `Date` |
| 5 | `KPR_Dates_EndOfMonth(DateIn As Variant) As Variant` | `Date` |
| 6 | `KPR_Dates_BeginOfQuarter(DateIn As Variant) As Variant` | `Date` |
| 7 | `KPR_Dates_EndOfQuarter(DateIn As Variant) As Variant` | `Date` |
| 8 | `KPR_Dates_BeginOfYear(DateIn As Variant) As Variant` | `Date` |
| 9 | `KPR_Dates_EndOfYear(DateIn As Variant) As Variant` | `Date` |
| 10 | `KPR_Dates_IsMonthEnd(DateIn As Variant) As Variant` | `Boolean` |
| 11 | `KPR_Dates_IsQuarterEnd(DateIn As Variant) As Variant` | `Boolean` |
| 12 | `KPR_Dates_IsYearEnd(DateIn As Variant) As Variant` | `Boolean` |
| 13 | `KPR_Dates_IsLeapYear(DateIn As Variant) As Variant` | `Boolean` |
| 14 | `KPR_Dates_AddDays(DateIn As Variant, nDays As Variant) As Variant` | `Date` |
| 15 | `KPR_Dates_AddWeeks(DateIn As Variant, nWeeks As Variant) As Variant` | `Date` |
| 16 | `KPR_Dates_AddMonths(DateIn As Variant, nMonths As Variant, Optional Opt_KeepEOM As Variant = False) As Variant` | `Date` |
| 17 | `KPR_Dates_AddYears(DateIn As Variant, nYears As Variant, Optional Opt_KeepEOM As Variant = False) As Variant` | `Date` |
| 18 | `KPR_Dates_NthWeekdayOfMonth(YearIn As Variant, MonthIn As Variant, WdIndex As Variant, n As Variant, Optional Opt_WeekBaseMonday As Variant = True) As Variant` | `Date` |
| 19 | `KPR_Dates_LastWeekdayOfMonth(YearIn As Variant, MonthIn As Variant, WdIndex As Variant, Optional Opt_WeekBaseMonday As Variant = True) As Variant` | `Date` |
| 20 | `KPR_Dates_PillarFromDates(StartDate As Variant, EndDate As Variant, Optional Opt_Rounding As Variant = "NEAREST") As Variant` | `String` |
| 21 | `KPR_Dates_DateFromPillar(StartDate As Variant, Pillar As Variant) As Variant` | `Date` |
| 22 | `KPR_Dates_HostDateSystem() As Variant` | `Long` 1900 or 1904 |

The first 21 functions are scalar/array-capable. `KPR_Dates_HostDateSystem()` remains scalar because it has no value argument. `KPR_Dates_DateFromPillar` replaces the current plural `KPR_Dates_DatesFromPillar`; this pre-release retains no compatibility alias.

### Value and shape rules

- Value arguments may be scalar, a `1×1` wrapper, a Range or a one-/two-dimensional in-memory array.
- A scalar or `1×1` expands to the target shape. Multiple non-scalar value arguments must have exactly the same row and column dimensions.
- Row, column and rectangular orientation is preserved. A one-dimensional VBA array is interpreted as `1×N`.
- There is no row/column outer product, implicit cross-broadcasting or `UsedRange` trimming.
- A scalar target returns a scalar Variant, not a `1×1` array. A non-scalar target returns a two-dimensional Variant array for Excel to spill.
- Element validation and incoming errors are isolated per element. A target-shape mismatch, invalid non-scalar `Opt_` argument, excessive target size or unavailable worksheet host is a call-level result.
- Blank required values and `Empty` are contract-invalid. `Null`, Boolean and disallowed object inputs are rejected. Numeric-looking strings are never reinterpreted as serials.

### Caller and date-system rules

`Application.Caller` is used only to distinguish the contexts that actually carry serial-date ambiguity:

1. If the caller is a `Range`, the implementation reads `Application.Caller.Parent.Parent.Date1904` from that exact workbook. A 1904 workbook returns call-level `#N/A` before element conversion.
2. If the caller is not a `Range`, the call is from VBA or infrastructure. The direct caller owns the meaning of values it constructed, and the library applies the documented 1900 serial contract.
3. `ActiveWorkbook`, `ThisWorkbook` and other unrelated workbooks are never used as fallback date-system authorities.

`KPR_Dates_HostDateSystem()` reports `1900` or `1904` for an identifiable worksheet caller and `1900` for a direct VBA/infrastructure call. A result that cannot be established in a worksheet host remains `#N/A`.

### Dynamic-array policy

The one public surface does not make scalar support depend on spill support. Scalar calls are tested on every Excel version for which scalar compatibility is claimed. Multi-cell calls are tested and claimed only on dynamic-array Excel. v0.0.2 makes no Ctrl+Shift+Enter compatibility claim, performs no version detection, and does not manufacture `#SPILL!`.

### Pillar policy

- Absolute intervals shorter than seven days remain exact day pillars.
- Whole-week and calendar-month anchors are computed from the start date.
- `FLOOR` chooses the latest non-overshooting anchor; `CEILING` chooses the earliest non-undershooting anchor.
- `NEAREST` chooses the closest anchor, with month representation preferred for equivalent anchors.
- Negative intervals round their absolute magnitude and then restore the sign.
- Rounded pillar conversion is deliberately not a general round-trip invariant.

## Production architecture

```text
src/modules/
  KPR_Core_Err.bas
  KPR_Core_Parse.bas
  KPR_Core_Dates.bas
  KPR_Core_Array.bas
  KPR_Dates.bas
  KPR_Register.bas
  KPR_UI_Bars.bas
  KPR_UI_Ribbon.bas
src/ribbon/
  customUI14.xml
```

| Module | Responsibility | Allowed production dependencies |
| --- | --- | --- |
| `KPR_Core_Err` | Native error construction, classification and propagation | None |
| `KPR_Core_Dates` | Pure Gregorian calculations and pillar core | None |
| `KPR_Core_Parse` | Strict scalar parsing, integer/control parsing and caller-date-system classification | `KPR_Core_Err` |
| `KPR_Core_Array` | Shape discovery, scalar expansion, exact-shape broadcasting and element traversal | `KPR_Core_Err` |
| `KPR_Dates` | The only supported public date facade; scalar dispatch and array traversal over shared scalar cores | `KPR_Core_Err`, `KPR_Core_Parse`, `KPR_Core_Dates`, `KPR_Core_Array` |
| `KPR_Register` | MacroOptions manifest and repeatable registration | Supported API names only; no date algorithms |
| `KPR_UI_Bars` | Temporary classic CommandBars build/teardown | Registration, demo and development test entry points |
| `KPR_UI_Ribbon` | Ribbon callbacks and invalidation | Registration, demo and development test entry points |

All `KPR_Core_*` modules use `Option Private Module`. Cross-module helpers may be technically `Public` only where VBA requires it; that does not make them supported API. Core modules never depend on the facade, registration, UI, tests or demo code. Tests and the demo consume the supported API rather than private helpers. No calendar or empty placeholder module is created.

## Registration and Excel UI

- `Application.MacroOptions` uses one category only: `KPR Dates`.
- One manifest covers exactly the 22 supported names, function descriptions and argument descriptions.
- Argument-description arrays are one-based, have exactly the signature's argument count, contain no blank entries and satisfy Excel's length constraints.
- Registration is safe to repeat and callable manually, during add-in startup and from both UI surfaces.
- RibbonX uses `customUI14.xml`, stable KPR IDs, exact callbacks and idempotent package injection into untracked Office artifacts.
- CommandBars use stable KPR ownership tags, remove stale KPR controls before rebuilding, and tolerate repeated teardown.
- Release UI exposes registration and demo actions. Regression actions, if present, are development-only.

## Regression architecture

```text
test/modules/
  KPR_Test_Runner.bas
  KPR_Test_Assert.bas
  KPR_Test_Fixtures_Generated.bas
  KPR_Test_Dates.bas
  KPR_Test_Shape.bas
  KPR_Test_State.bas
  KPR_Test_Oracle.bas
test/fixtures/
  date-fixtures.tsv
test/evidence/
  README.md
  evidence-schema.json
  compile-template.md
tools/
  gen_fixtures.py
```

`date-fixtures.tsv` is the canonical generated fixture artifact. `KPR_Test_Fixtures_Generated.bas` is generated test source, not hand-maintained production code. The independent deterministic generator has a drift-check mode and records enough source metadata for reproducibility.

The runner interface is:

```vb
Public Function KPR_Test_RunAll( _
    ByVal SourceSha As String, _
    ByVal OutputFolder As String) As Boolean
```

The runner is callable directly and through `Application.Run`, preserves caller state, returns Boolean success, and writes a structured summary, case-level TSV and environment record for the supplied exact source SHA. Direct VBA calls exercise the documented 1900 caller contract; worksheet-host tests separately exercise 1900 and 1904 behavior.

Coverage includes accepted and rejected inputs, serial limits and serial 60, leap years, month/quarter/year boundaries, `AddDays`/`AddWeeks`/`AddMonths`/`AddYears`, pillar grammar and all rounding modes, both weekday bases, scalar/array parity, row/column/rectangular and one-dimensional shapes, mismatches, optional-argument rejection, blanks, mixed validity, native errors, the 100,000-element boundary, overflow, repeatability, caller-state restoration and date-system behavior.

`KPR_Test_Oracle.bas` owns allowed worksheet-oracle checks against `EOMONTH`, `EDATE`, `WEEKDAY`, `DAY`, `YEAR` and `MONTH` only where contracts overlap. `WORKDAY.INTL` and `NETWORKDAYS.INTL` are out of scope.

Schemas and templates are tracked under `test/evidence/`. Actual runs are written under ignored `test-results/` and attached to the final certification issue and pre-release. Evidence identifies the exact tested SHA; it is not committed back into that same candidate and therefore does not create an evidence/SHA recursion.

## Demo strategy

`demo/modules/KPR_Demo_Dates.bas` deterministically builds the demonstration workbook from a clean Excel instance and an explicit output path. The builder is authoritative source. Generated `.xlsx`, `.xlsm`, `.xlam` and other Office binaries are never committed; a certified demo or add-in may be attached later as a release asset.

The demo uses the single public function surface for both scalar and array examples. It shows supported behavior, native errors and array shapes without claiming calendar support, business-day support, production readiness, legacy CSE use or any untested Excel version.

## Delivery phases

| Phase | Outcome | Issues |
| --- | --- | --- |
| 1. Contract and export baseline | Normative behavior and importable VBE source format | #9–#10 |
| 2. Core and public date layer | Layered modules, strict parsing, caller/date-system policy, pillar policy and 22 functions | #11–#15 |
| 3. Array engine and registration | Shared shape engine, array-capable public facade and 22-entry MacroOptions manifest | #16–#18 |
| 4. Regression | Independent fixtures, runner, full suites and allowed Excel oracles | #19–#22 |
| 5. Demo and UI | Deterministic demo, RibbonX and classic CommandBars | #23–#25 |
| 6. Surface freeze and candidate | API classification, expanded static checks, docs, VERSION and CHANGELOG | #26–#28 |
| 7. Certification and release | Exact-source Windows import, compile, regression evidence and pre-release | #29 |

Parallel work is allowed only where the issue dependency list permits it. The `blocked` label remains on a dependent issue until its prerequisites are satisfied.

## Exit gate

- The 21 v0.0.2 issues are complete, or issue #29 is the sole remaining issue while exact-source evidence is assembled.
- The 22-name supported surface, MacroOptions manifest, documentation and static inventory agree exactly; no duplicate spill API or calendar placeholder exists.
- Hosted static checks pass at the exact candidate SHA without claiming Excel execution.
- A clean Windows Excel import and VBA compilation succeed at that same SHA.
- Scalar regressions pass from direct VBA under the documented 1900 contract and from worksheet cells in a 1900 workbook.
- Multi-cell regression and scalar/array parity pass on dynamic-array Excel; no CSE compatibility is claimed.
- A 1904 worksheet call is refused with call-level `#N/A`, and `HostDateSystem` reports the documented caller context.
- Allowed cross-oracle tests, UI lifecycle tests, repeatability tests and caller-state restoration pass.
- The demo is rebuilt from source and its claims match the tested evidence.
- `VERSION` is `0.0.2`; README and CHANGELOG explicitly defer calendars to v0.0.3 and business-day arithmetic to v0.0.4.
- Generated Office binaries remain absent from git; `v0.0.1` remains unchanged.
- Evidence is attached to issue #29 and the pre-release, identifying the exact tested SHA and Excel/Windows environment.
- Only after every gate passes may the protected `v0.0.2` tag and GitHub pre-release be created from the certified SHA.

## Milestone issue register

This terminal register is generated from the live GitHub issues after the architecture reconciliation. It is intentionally complete: every issue includes its current title, state, assignee, labels, URL and full body. Because the register must remain the end of this document, issue #29 is the final content.

<details>
<summary><strong>#9 — Specify the date-layer behavioural contract</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `documentation`, `behavior-change`, `P1`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/9

#### Body

## Objective

Freeze the complete supported contract for the v0.0.2 date layer before production code is reorganized.

## Decisions to record

- Publish one supported public surface containing exactly 22 `KPR_Dates_*` names. Every value-taking function is array-capable; scalar input returns a scalar and multi-cell input returns a shape-preserving array. Do not create `_Spill` twins or `KPR_Dates_Spill.bas`.
- Keep `KPR_Dates_HostDateSystem()` scalar because it has no value argument.
- Rename the current plural `KPR_Dates_DatesFromPillar` member to `KPR_Dates_DateFromPillar`; do not retain a compatibility alias in this pre-release.
- Put required value arguments first and optional `Opt_` arguments last. Value arguments may vectorize; `Opt_` arguments must be omitted, scalar or 1×1.
- Accept locale-independent ISO `YYYY-MM-DD` text only; reject locale-formatted and numeric-looking strings.
- Retain the supported window `1900-03-01 .. 9999-12-31`, excluding the Excel serial-60 anomaly and all earlier serials.
- Distinguish worksheet and VBA callers: a worksheet `Range` caller is checked for `Date1904` and a 1904 host returns `#N/A`; a non-`Range` caller is a direct VBA/infrastructure call and uses the documented 1900 serial contract. Never fall back to an unrelated active workbook.
- Define native errors in three categories: `#VALUE!` for uninterpretable/contract-invalid input, `#NUM!` for validly formed values outside the supported domain, and `#N/A` when the result is unavailable in the host configuration. Propagate incoming Excel errors verbatim.
- Define scalar expansion, exact-shape broadcasting, orientation preservation, one-dimensional VBA arrays as 1×N, blank handling, per-element errors and call-level shape failures.
- Cap a single array-capable call at 100,000 output elements; 100,001 or more returns call-level `#NUM!`. Never shorten a supplied Range through `UsedRange`.
- Support scalar calls on every Excel version ultimately certified for the scalar surface. Support and claim multi-cell use only on dynamic-array Excel. Make no Ctrl+Shift+Enter compatibility claim until separately tested.
- Define the `NEAREST`, `FLOOR` and `CEILING` pillar policies and make `NEAREST` the default.
- Distinguish supported public API from procedures that must be technically `Public` for cross-module VBA, RibbonX, CommandBars, `Application.Run` or testing.
- Reserve `KPR_Cal_*` for v0.0.3. Do not add calendar options to pure `KPR_Dates_*` functions.

## Deliverable

Add a source-controlled date-layer contract document that is the normative reference for implementation, registration, tests and documentation.

## Acceptance criteria

- [ ] Every one of the 22 supported names has an exact VBA signature, semantic return type and vectorization classification.
- [ ] The input-type and native-error matrices contain no locale-dependent parsing path.
- [ ] The `1900-03-01 .. 9999-12-31` gate and worksheet-versus-VBA caller policy are unambiguous.
- [ ] `#VALUE!`, `#NUM!`, `#N/A` and verbatim error propagation have distinct documented meanings.
- [ ] Scalar expansion, optional-argument rules, orientation, 1-D arrays, blanks, the 100,000-element cap and per-element error propagation are specified.
- [ ] Scalar and dynamic-array compatibility claims are stated separately with no CSE claim.
- [ ] Pillar rounding is defined with tie-breaking, negative intervals and non-invariant round trips.
- [ ] The supported-API boundary and reserved `KPR_Cal_*` namespace are recorded.
- [ ] Calendars, weekend masks, holidays, business-day arithmetic and roll conventions are explicitly out of scope.

## Dependencies

None. This issue is the normative dependency for all behavior-changing implementation work.

</details>

<details>
<summary><strong>#10 — Adopt VBE export format and assert `Attribute VB_Name`</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `ci`, `P2`, `repository`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/10

#### Body

## Objective

Make every tracked VBA module a deterministic, importable VBE export rather than a hand-written approximation of one.

## Scope

- Add the exported `Attribute VB_Name = "…"` header to every tracked `.bas` module.
- Require the attribute value to match the filename and intended VBA component name exactly.
- Preserve `Option Explicit`, CRLF working-tree exports and the repository's existing source-first encoding policy.
- Extend the static checker with positive and negative fixtures for missing, mismatched and duplicate module names.
- Document the export/import procedure contributors must follow.

## Acceptance criteria

- [ ] Every production, test and demo `.bas` file is a valid VBE-style export with a unique matching `Attribute VB_Name`.
- [ ] Static checks reject a missing header, a mismatched filename/module name and duplicate module names.
- [ ] The checker continues to require `Option Explicit`.
- [ ] A documented Windows VBE export/import round trip is reserved for final certification.
- [ ] This issue makes no Excel compilation or execution claim.

## Dependencies

None.

</details>

<details>
<summary><strong>#11 — Split `KPR_Dates_Days` into the layered architecture</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `blocked`, `code`, `P1`, `refactor`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/11

#### Body

## Objective

Replace the monolithic `KPR_Dates_Days.bas` baseline with explicit core, public facade, registration and UI boundaries without creating a duplicated array API.

## Target production inventory

- `KPR_Core_Err.bas` — native error construction, classification and propagation helpers.
- `KPR_Core_Parse.bas` — strict date, integer and ISO parsing.
- `KPR_Core_Dates.bas` — date arithmetic, boundary logic and pillar core.
- `KPR_Core_Array.bas` — shape discovery, scalar expansion, broadcast resolution and allocation.
- `KPR_Dates.bas` — the complete supported 22-name scalar/array-capable date API.
- `KPR_Register.bas` — MacroOptions manifest and registration.
- `KPR_UI_Bars.bas` — classic CommandBars lifecycle.
- `KPR_UI_Ribbon.bas` — Ribbon callbacks only.

Do not create `KPR_Dates_Spill.bas`, `KPR_Cal.bas` or `KPR_Core_Cal.bas`.

## Visibility and dependency rule

Use `Option Private Module` for internal core modules. Cross-module procedures may be technically `Public` only where VBA requires it; all other helpers remain `Private`. Technical visibility never makes a member supported API.

`KPR_Dates` may call `KPR_Core_Parse`, `KPR_Core_Dates`, `KPR_Core_Array` and `KPR_Core_Err`. Core modules must not depend on the public facade, registration, UI, tests or demo code. Registration and UI modules contain no date algorithms.

## Acceptance criteria

- [ ] Existing calculations are migrated without leaving duplicate implementations.
- [ ] All 22 supported functions exist only in `KPR_Dates.bas`.
- [ ] Scalar and array paths share one element implementation per function.
- [ ] Internal modules use `Option Private Module` and the narrowest practical member visibility.
- [ ] Registration, callbacks, tests and demo builders contain no date algorithms.
- [ ] The old `KPR_Dates_Days.bas` file is removed after migration.
- [ ] No spill or calendar placeholder module is created.
- [ ] The module ownership and allowed-dependency matrix is documented.

## Dependencies

- [ ] #9
- [ ] #10

</details>

<details>
<summary><strong>#12 — Implement strict date and integer parsing with error propagation</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `behavior-change`, `blocked`, `code`, `P1`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/12

#### Body

## Objective

Implement the input contract from #9 once in the core parsing layer so scalar and multi-cell evaluation cannot diverge.

## Scope

- Parse ISO text only in exact `YYYY-MM-DD` form and validate components without DateSerial rollover.
- Reject locale-formatted dates, numeric-looking strings and permissive conversions.
- Accept native VBA `Date` and numeric 1900-system serials only inside `1900-03-01 .. 9999-12-31`.
- Reject Empty, Null, Boolean, blank required cells, non-Range objects and unsupported array wrappers intentionally.
- Parse integer arguments without silent truncation, Boolean coercion, banker's rounding or overflow.
- Propagate incoming `CVErr` values unchanged.
- Return `#VALUE!` for uninterpretable/contract-invalid inputs and `#NUM!` for well-formed values outside the supported domain.
- Keep parsing free of UI, worksheet selection, active-workbook and locale state.

The worksheet/VBA caller distinction and 1904 host refusal are owned by #13, outside the scalar parser.

## Acceptance criteria

- [ ] No production path relies on locale-sensitive `CDate`, `IsDate` or equivalent permissive parsing.
- [ ] ISO leap-day and component validation is exact.
- [ ] The minimum and maximum supported dates are accepted and adjacent out-of-window values return `#NUM!`.
- [ ] Integer-required arguments reject fractions and Boolean values with `#VALUE!`; out-of-Long values return `#NUM!`.
- [ ] Incoming native Excel errors propagate verbatim.
- [ ] Invalid input never yields a plausible date, a message string, a `MsgBox` or an unhandled VBA error.
- [ ] Deterministic fixtures cover every accepted and rejected input class.

## Dependencies

- [ ] #11

</details>

<details>
<summary><strong>#13 — Enforce the worksheet/VBA date-system policy and add `HostDateSystem`</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `documentation`, `behavior-change`, `blocked`, `code`, `P1`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/13

#### Body

## Objective

Prevent 1,462-day worksheet serial shifts while preserving direct VBA use and the regression harness.

## Caller contract

- Inspect `Application.Caller` once at the public-call boundary, before any element loop.
- If the caller is a worksheet `Range`, read that Range's workbook `Date1904` setting.
- A 1900 worksheet host proceeds normally.
- A 1904 worksheet host returns one call-level `#N/A`, meaning the result is unavailable in this host configuration.
- If the caller is not a `Range`, treat the call as direct VBA/infrastructure use and assume the documented 1900 serial contract. The VBA caller owns the interpretation of values it constructed.
- Never use `ActiveWorkbook`, `ThisWorkbook` or another unrelated workbook as a fallback.

## Diagnostic function

Add scalar-only `KPR_Dates_HostDateSystem()`:

- return `1900` or `1904` for an identifiable worksheet caller;
- return `1900` for a direct VBA/infrastructure caller under the documented VBA contract;
- return `#N/A` only when a worksheet host should exist but cannot be resolved reliably.

## Error taxonomy

- `#VALUE!` — uninterpretable or contract-invalid input;
- `#NUM!` — well-formed value outside the supported date/numerical domain;
- `#N/A` — result unavailable in the current host configuration;
- incoming native errors — propagated unchanged.

## Acceptance criteria

- [ ] Every public worksheet call performs the date-system guard once, not once per array element.
- [ ] A 1904 worksheet call cannot return a plausible shifted scalar or array.
- [ ] Direct VBA calls and `KPR_Test_RunAll` execute under the documented 1900 contract.
- [ ] `HostDateSystem` returns the documented result for 1900 worksheet, 1904 worksheet and VBA callers.
- [ ] No active-workbook fallback exists.
- [ ] The three native-error categories remain distinct in scalar and array cases.
- [ ] Windows certification exercises 1900 and 1904 worksheets plus direct VBA calls.

## Dependencies

- [ ] #12

</details>

<details>
<summary><strong>#14 — Define and implement the three pillar-rounding modes</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `behavior-change`, `blocked`, `code`, `P2`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/14

#### Body

## Objective

Replace the unresolved nearest-versus-floor behavior with one explicit pillar conversion policy shared by scalar and array calls through the single public surface.

## Contract

- Add optional `Opt_Rounding` after all value arguments.
- Accept `NEAREST` (default), `FLOOR` and `CEILING`, case-insensitively; reject every other value.
- Preserve exact day pillars for absolute intervals shorter than seven days.
- Compare valid whole-week and calendar-month anchors from the start date.
- `FLOOR` selects the latest non-overshooting anchor; `CEILING` selects the earliest non-undershooting anchor; `NEAREST` selects the closest anchor.
- Prefer the month representation when week and month candidates land on the same date.
- For negative intervals, round the absolute magnitude and then restore the sign.
- Document that `DateFromPillar(PillarFromDates(...))` is not a general invariant after rounding.

## Acceptance criteria

- [ ] Scalar and array calls reach one pillar core through the same public functions.
- [ ] All three modes cover positive, negative, exact, tie and boundary cases.
- [ ] Month-end and leap-day anchors are deterministic.
- [ ] Invalid modes return the contract's native error.
- [ ] Formatting and parsing are mutually consistent for every supported pillar token.
- [ ] Regression fixtures explicitly demonstrate non-invariant rounded round trips.
- [ ] No duplicate pillar UDF or `_Spill` twin is created.

## Dependencies

- [ ] #11
- [ ] #12

</details>

<details>
<summary><strong>#15 — Complete the 22-function date primitive surface</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `enhancement`, `blocked`, `code`, `P2`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/15

#### Body

## Objective

Deliver the complete element-correct 22-name public date surface in `KPR_Dates.bas`, ready for the array loop added by #17.

## Supported surface

Existing/re-signatured members: `DayOfWeek`, `DaysInMonth`, `DaysInYear`, `BeginOfMonth`, `EndOfMonth`, `IsMonthEnd`, `IsQuarterEnd`, `IsYearEnd`, `IsLeapYear`, `AddWeeks`, `AddMonths`, `AddYears`, `NthWeekdayOfMonth`, `LastWeekdayOfMonth`, `PillarFromDates`, and singular `DateFromPillar`.

New members: `AddDays`, `BeginOfQuarter`, `EndOfQuarter`, `BeginOfYear`, `EndOfYear`, and `HostDateSystem`.

## Acceptance criteria

- [ ] All 22 functions use the exact signatures in #9 and return `Variant` where native errors are possible.
- [ ] Month, quarter and year boundaries are correct across leap years and year transitions.
- [ ] `AddDays`, `AddWeeks`, `AddMonths` and `AddYears` implement the documented clipping/preservation and overflow rules.
- [ ] Weekday bases and nth/last weekday locators reject unsupported arguments intentionally.
- [ ] Pillar functions use the policy from #14 and expose only the singular `DateFromPillar` name.
- [ ] Every value-taking function has one element implementation usable by scalar and array calls.
- [ ] `HostDateSystem` implements the caller policy from #13 and remains scalar.
- [ ] No `_Spill` name, calendar, holiday, weekend-mask, business-day or roll-convention function is introduced.

## Dependencies

- [ ] #12
- [ ] #13
- [ ] #14

</details>

<details>
<summary><strong>#16 — Implement the array shape and broadcasting engine</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `blocked`, `code`, `P1`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/16

#### Body

## Objective

Create one internal array engine with deterministic orientation, broadcasting, optional-argument and capacity semantics for the single public surface.

## Contract

- Value arguments may vectorize. `Opt_` arguments must be omitted, scalar or 1×1; a larger optional argument returns call-level `#VALUE!`.
- A scalar or 1×1 value expands to the resolved output shape.
- All non-scalar value arguments must have identical row and column dimensions.
- Do not perform row-to-column outer products or implicit cross-broadcasting.
- An all-scalar call returns a scalar, not a 1×1 array.
- Preserve worksheet Range orientation. Treat a one-dimensional VBA array as 1×N.
- Reject multi-area, empty, jagged and unsupported array inputs intentionally.
- A blank required element produces `#VALUE!` at that output position.
- An error element propagates its exact native error without suppressing valid neighbors.
- A shape conflict returns one call-level `#VALUE!`.
- Cap the resolved output at 100,000 elements. A larger call returns one call-level `#NUM!`.
- Never intersect a supplied Range with `UsedRange` or silently shorten the requested output.
- Own shape classification, broadcast resolution, allocation and 1×1 unwrapping only; contain no date algorithms and no generic function-pointer dispatch.

## Acceptance criteria

- [ ] Scalar, 1×1, row, column and rectangular fixtures preserve the documented result type and shape.
- [ ] Multi-argument scalar expansion works identically for Ranges and in-memory arrays.
- [ ] One-dimensional VBA arrays are asserted as 1×N.
- [ ] Non-scalar optional arguments, shape mismatches and unsupported arrays return the documented call-level error.
- [ ] Exactly 100,000 elements are accepted and 100,001 returns `#NUM!`.
- [ ] Mixed-validity arrays preserve valid values and per-element errors.
- [ ] Traversal is deterministic and does not select, activate or recalculate unrelated Excel state.
- [ ] Caller calculation, events, screen updating and selection state remain unchanged.

## Dependencies

- [ ] #11
- [ ] #12

</details>

<details>
<summary><strong>#17 — Make the public date surface array-capable</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `enhancement`, `behavior-change`, `blocked`, `code`, `P1`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/17

#### Body

## Objective

Make the single 22-name `KPR_Dates_*` surface handle both scalar and multi-cell value inputs without introducing duplicated function names or modules.

## Implementation contract

- Every value-taking public function uses the common shape services from #16 and calls its own shared element implementation.
- Scalar evaluation is the 1×1 case of the same path and returns a scalar.
- Multi-cell evaluation returns a two-dimensional Variant array with the resolved shape.
- `NthWeekdayOfMonth` and `LastWeekdayOfMonth` vectorize like the other value-taking functions.
- `HostDateSystem` remains scalar because it has no value argument.
- Blanks and native errors follow the per-element policy; shape conflicts, non-scalar `Opt_` arguments and the element cap fail at call level.
- The date-system guard from #13 runs once per public call before array traversal.
- Do not create `_Spill` twins, `KPR_Dates_Spill.bas` or parallel scalar/array algorithms.

## Compatibility policy

- Scalar calls are supported on each Excel version ultimately certified for scalar use.
- Multi-cell calls are tested, supported and claimed only on dynamic-array Excel.
- Make no Ctrl+Shift+Enter or other legacy multi-cell compatibility claim until separately tested.
- Do not add version detection or attempt to manufacture Excel's spill-placement errors.

## Acceptance criteria

- [ ] Exactly 22 supported public names remain.
- [ ] Every value-taking function returns the same value/error for a scalar, a 1×1 wrapper and the corresponding element of an N-element call.
- [ ] Row, column and rectangle orientation follows #16.
- [ ] The 100,000-element cap and optional-argument rule apply uniformly.
- [ ] No duplicate public calculation function or spill module exists.
- [ ] Documentation distinguishes broad scalar support from dynamic-array-only multi-cell support.

## Dependencies

- [ ] #15
- [ ] #16

</details>

<details>
<summary><strong>#18 — Add the MacroOptions manifest and registration lifecycle</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `documentation`, `enhancement`, `blocked`, `code`, `P2`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/18

#### Body

## Objective

Register the single supported worksheet surface consistently without mixing descriptions and lifecycle code into calculation modules.

## Scope

- Maintain one manifest for exactly 22 supported functions, their category, descriptions and argument descriptions.
- Use exactly one MacroOptions category: `KPR Dates`.
- Describe both scalar behavior and dynamic-array-only multi-cell behavior for value-taking functions.
- Build `ArgumentDescriptions` as a 1-based array whose length exactly matches the signature, contains no blank element and respects Excel's supported description length.
- Add explicit, repeatable register-all and best-effort unregister entry points suitable for manual use, add-in startup and UI delegation.
- Registration overwrites by function name and is idempotent by construction.
- Keep infrastructure procedures technically public only where Excel requires it and classify them as unsupported infrastructure.
- Preserve caller/application state and avoid selections, activation, alerts or duplicate registration side effects.
- Keep calendar and business-day categories out of v0.0.2.

## Acceptance criteria

- [ ] Every one of the 22 supported functions has exactly one manifest record.
- [ ] No `_Spill` name or infrastructure procedure appears in the supported manifest.
- [ ] Function and argument descriptions match the normative signatures and optional-argument order.
- [ ] Argument-description arrays are 1-based, complete, nonblank and within the documented length limit.
- [ ] Repeated registration produces the same MacroOptions state and no duplicate UI artifact.
- [ ] Repeated best-effort unregistration is controlled and documented.
- [ ] Registration can run through `Application.Run` under the direct-VBA 1900 contract.
- [ ] Manifest completeness and uniqueness can be checked statically.
- [ ] Only `KPR Dates` is registered.

## Dependencies

- [ ] #15
- [ ] #17

</details>

<details>
<summary><strong>#19 — Add the independent fixture generator and generated fixture module</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `blocked`, `ci`, `P1`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/19

#### Body

## Objective

Generate deterministic expected results independently of the VBA implementation and make the complete scalar/array contract reviewable in source control.

## Deliverables

- `tools/gen_fixtures.py`
- `test/fixtures/date-fixtures.tsv`
- `test/modules/KPR_Test_Fixtures_Generated.bas`
- generated-source classification in `.gitattributes`

## Fixture contract

Each case must have a stable ID, suite, input-kind metadata, arguments, expected result type, expected value or native error, and a short rationale. The generated VBA module is derived from the canonical TSV and must not be hand-edited.

The generator may use Python standard-library date primitives, but must not invoke KPR code or mechanically translate the same VBA algorithm line for line.

## Acceptance criteria

- [ ] Generation is deterministic across repeated runs and platforms.
- [ ] A `--check` mode fails when committed TSV or VBA output is stale.
- [ ] The generated module is classified as generated source and remains a valid VBE export with `Option Explicit`.
- [ ] Fixtures cover accepted/rejected inputs, `1900-03-01 .. 9999-12-31`, leap years, date boundaries, arithmetic, pillar modes and weekday bases.
- [ ] Shape fixtures cover scalar, 1×1, row, column, rectangle, 1-D VBA arrays, optional-argument rejection and shape mismatch.
- [ ] Capacity fixtures cover exactly 100,000 and 100,001 elements.
- [ ] Expected `#VALUE!`, `#NUM!`, `#N/A` and propagated native errors are encoded explicitly.
- [ ] Fixture IDs and ordering remain stable unless a reviewed contract change requires an update.
- [ ] Calendar, weekend-mask, holiday and business-day fixtures are absent.

## Dependencies

- [ ] #12
- [ ] #14
- [ ] #15

</details>

<details>
<summary><strong>#20 — Add the regression runner, assertions and evidence schema</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `blocked`, `code`, `P1`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/20

#### Body

## Objective

Provide a deterministic, automation-friendly VBA test runner with structured exact-source evidence and complete caller-state restoration.

## Test modules

- `KPR_Test_Assert.bas`
- `KPR_Test_Runner.bas`
- committed evidence documentation and schema under `test/evidence/`

## Public test-infrastructure interface

- `KPR_Test_RunAll(ByVal SourceSha As String, ByVal OutputFolder As String) As Boolean`
- `KPR_Test_RunSuite(ByVal SuiteName As String, ByVal SourceSha As String, ByVal OutputFolder As String) As Boolean`

These entry points are callable through `Application.Run`, but are not supported production API. Calls from the runner into `KPR_Dates_*` are direct VBA calls and therefore use the documented 1900 serial contract rather than requiring a worksheet `Range` caller.

## Evidence contract

Write a machine-readable summary, a case-level TSV and environment metadata covering the exact source SHA, Excel version/build/bitness, Windows version, locale, workbook date system, caller context, suite/case totals, failures and elapsed time. Keep schemas/templates under `test/evidence/`; write actual runs to ignored `test-results/` for attachment to the final certification issue and pre-release.

## Acceptance criteria

- [ ] Assertions cover values, dates/serials, Booleans, strings, native error codes and array shape/content.
- [ ] A failed assertion records the case and continues unless the harness itself cannot proceed.
- [ ] Runs are deterministic apart from explicitly identified environment/time fields.
- [ ] The runner returns `True` only when every selected case passes.
- [ ] Direct VBA calls execute under the documented 1900 caller contract.
- [ ] Calculation mode, events, screen updating, alerts, status bar, active workbook/sheet and selection are restored after success and failure.
- [ ] No automated path displays a `MsgBox` or leaves files/workbooks open.
- [ ] Evidence validates against the committed schema and records the supplied exact source SHA.

## Dependencies

- [ ] #19

</details>

<details>
<summary><strong>#21 — Add contract, shape, parity, error and state regression suites</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `blocked`, `P1`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/21

#### Body

## Objective

Exercise the complete single public date surface and prove scalar/array parity without relying on visual inspection.

## Suites

- `KPR_Test_Dates.bas`
- `KPR_Test_Shape.bas`
- `KPR_Test_State.bas`

## Required coverage

Accepted and rejected date inputs; minimum/maximum serial boundaries; leap years and 29 February; month, quarter and year boundaries; `AddDays`, `AddWeeks`, `AddMonths` and `AddYears`; clipping/preservation rules; pillar parse/format and all three rounding modes; non-invariant rounded round trips; weekday bases and locators; all 22 public names; scalar, 1×1, row, column and rectangular inputs; 1-D VBA arrays as 1×N; scalar expansion; non-scalar `Opt_` rejection; shape mismatches; blanks; mixed validity; native errors; exactly 100,000 versus 100,001 elements; overflow; repeatability; state restoration; worksheet/VBA caller distinction; `HostDateSystem`; and 1904 worksheet refusal.

## Acceptance criteria

- [ ] Every supported function has positive, edge and invalid-domain cases.
- [ ] Every value-taking function is compared across scalar, 1×1 and corresponding array elements.
- [ ] Row, column and rectangle orientation is asserted, not inferred from displayed values.
- [ ] Mixed arrays prove that one invalid/error element does not corrupt valid neighbors.
- [ ] `#VALUE!`, `#NUM!`, `#N/A` and propagated input errors remain distinct.
- [ ] Direct VBA tests run under the 1900 contract without a worksheet caller.
- [ ] 1900 worksheet calls succeed and 1904 worksheet calls return one call-level `#N/A`.
- [ ] State restoration is tested after both passing and deliberately failing cases.
- [ ] Two consecutive runs produce identical case results.
- [ ] The suite contains no `_Spill` name, calendar, weekend-mask, holiday or business-day arithmetic case.

## Dependencies

- [ ] #17
- [ ] #20

</details>

<details>
<summary><strong>#22 — Add Excel cross-oracle checks for the date surface</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `blocked`, `P2`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/22

#### Body

## Objective

Compare KPR date primitives with independent native Excel worksheet functions where their contracts genuinely overlap.

## Deliverable

Add `test/modules/KPR_Test_Oracle.bas`.

## Oracle scope

Use only `EOMONTH`, `EDATE`, `WEEKDAY`, `DAY`, `YEAR` and `MONTH`. Construct comparisons only for scalar/element inputs where Excel and KPR have matching documented semantics.

Do not treat Excel's permissive parsing, implicit coercion, pre-1900-03-01 serial behavior or 1904 interpretation as authoritative when it differs from the KPR contract.

## Acceptance criteria

- [ ] Oracle cases state the exact overlap assumption being tested.
- [ ] Boundary and randomized deterministic samples cover month ends, leap years, arithmetic and weekday indices.
- [ ] A fixed seed and case order make repeated runs reproducible.
- [ ] Expected exclusions are documented rather than silently skipped.
- [ ] Oracle failures appear in the standard evidence output.
- [ ] The module is present in the documented and statically checked test inventory.
- [ ] `WORKDAY.INTL` and `NETWORKDAYS.INTL` are not used in v0.0.2.

## Dependencies

- [ ] #15
- [ ] #20

</details>

<details>
<summary><strong>#23 — Add the deterministic date-demo builder</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `documentation`, `enhancement`, `blocked`, `code`, `P3`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/23

#### Body

## Objective

Provide a reproducible demonstration workbook for the single scalar/array-capable surface without committing an Office binary or making unsupported claims.

## Deliverable

Add `demo/modules/KPR_Demo_Dates.bas` with a callable builder that creates a new workbook from a clean Excel instance, lays out fixed examples, applies deterministic formatting and saves only to an explicit output path.

The tracked builder is authoritative. A generated workbook may be attached later as a v0.0.2 release asset after exact-source certification.

## Acceptance criteria

- [ ] The builder creates the same sheets, labels, formulas, examples, widths, styles and named areas on repeated runs.
- [ ] Examples use the same 22 public names for scalar and multi-cell calls; no `_Spill` function is shown.
- [ ] Scalar examples are clearly distinguished from dynamic-array-only multi-cell examples.
- [ ] Examples cover native-error categories, array shapes, the 100,000-element policy and documented pillar behavior without creating an excessive workbook.
- [ ] The workbook labels calendars and business-day calculations as future scope.
- [ ] No example claims CSE compatibility, certified accuracy, production readiness or compatibility beyond the agreed evidence.
- [ ] The builder preserves Excel state and never overwrites an existing path without a controlled error.
- [ ] No `.xlsm`, `.xlam`, `.xlsx` or other generated Office binary is committed.
- [ ] Demo generation is exercised during Windows certification.

## Dependencies

- [ ] #17
- [ ] #18

</details>

<details>
<summary><strong>#24 — Add Ribbon callbacks, RibbonX and safe package injection</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `enhancement`, `blocked`, `code`, `P3`, `repository`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/24

#### Body

## Objective

Add a modern Excel Ribbon surface whose callbacks delegate to registration, demo and test infrastructure without becoming supported calculation API.

## Deliverables

- `src/modules/KPR_UI_Ribbon.bas`
- `src/ribbon/customUI14.xml`
- a deterministic, safe package-injection helper for untracked workbook/add-in artifacts

## Ribbon contract

Use the Office 2010+ `customUI/2009/07` schema, stable KPR IDs, built-in `imageMso` icons and exact callback signatures for `onLoad`, actions and dynamic enabled/visible state. The release surface provides registration and demo actions. Any regression action is development-only and must not appear in the normal release package.

## Acceptance criteria

- [ ] Callback names/signatures in RibbonX and VBA match exactly.
- [ ] `onLoad` caches the Ribbon object safely and invalidation is controlled.
- [ ] Repeated loads/actions do not create duplicate controls or duplicate registration.
- [ ] Callbacks contain no date algorithm and delegate to the owning module.
- [ ] Package injection validates the ZIP package, content types and relationships before changing an untracked target.
- [ ] Repeated injection replaces the KPR customUI part idempotently and creates no duplicate relationships.
- [ ] A failed injection leaves the original target recoverable and produces no committed binary.
- [ ] RibbonX parses in static checks and is loaded in Windows Excel certification.

## Dependencies

- [ ] #18
- [ ] #20
- [ ] #23

</details>

<details>
<summary><strong>#25 — Add idempotent classic CommandBars integration</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `enhancement`, `blocked`, `code`, `P3`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/25

#### Body

## Objective

Provide a dedicated classic Excel toolbar/menu lifecycle for environments and users that rely on CommandBars.

## Lifecycle

Expose technically public infrastructure entry points to install and remove a temporary KPR Dates bar/menu. Use a stable unique name and tags, remove stale KPR-owned controls before building, and bind actions through fully qualified `OnAction` targets.

The release bar provides registration and demo commands. A regression command may be present only in a development/test build.

## Acceptance criteria

- [ ] Install/build and teardown entry points are safe to call repeatedly in any order.
- [ ] A second install produces no duplicate bar, menu or control.
- [ ] Teardown removes only KPR-owned controls and tolerates already-absent UI.
- [ ] Temporary controls do not persist uncontrolled Excel state after the host closes.
- [ ] Callbacks delegate to registration, demo or test modules and contain no date calculations.
- [ ] Failures restore alerts, events, status bar, active workbook/sheet and selection.
- [ ] Windows certification covers install, repeat install, action dispatch, teardown and repeat teardown.

## Dependencies

- [ ] #18
- [ ] #20
- [ ] #23

</details>

<details>
<summary><strong>#26 — Freeze the public-API manifest and classify infrastructure members</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `documentation`, `blocked`, `P2`, `repository`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/26

#### Body

## Objective

Create one auditable inventory that separates the supported single public surface from internal and host-required procedures.

## Deliverable

Add a machine-checkable public-API manifest and readable documentation covering:

- exactly 22 supported `KPR_Dates_*` functions;
- the vectorization classification of each function, with `HostDateSystem` scalar-only;
- internal `KPR_Core_*` procedures protected by `Option Private Module`;
- MacroOptions registration entry points;
- Ribbon and CommandBars callbacks;
- regression-runner and demo-builder entry points.

Reserve `KPR_Cal_*` without creating placeholder modules or functions. Do not reserve or document a `KPR_Dates_*_Spill` surface.

## Acceptance criteria

- [ ] Exactly 22 functions are classified as supported v0.0.2 API.
- [ ] Every technically public infrastructure/internal member is explicitly marked unsupported.
- [ ] Names, signatures, vectorization and argument ordering agree with #9 and the MacroOptions manifest.
- [ ] No supported function is implemented in a core, UI, registration, test or demo module.
- [ ] No unclassified public procedure remains.
- [ ] No `_Spill` name or `KPR_Dates_Spill.bas` appears.
- [ ] The manifest is suitable for static completeness and drift checks.
- [ ] Future calendar/business-day namespaces and milestone boundaries are recorded.

## Dependencies

- [ ] #24
- [ ] #25

</details>

<details>
<summary><strong>#27 — Extend static checks for modules, manifests, fixtures and RibbonX</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `blocked`, `ci`, `P2`, `repository`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/27

#### Body

## Objective

Extend the existing repository-integrity gate to cover the final single-surface date architecture while keeping Excel execution claims separate.

## New checks

- Required production, test, fixture, RibbonX, demo and contract files exist.
- The production inventory contains the eight approved modules and no `KPR_Dates_Spill.bas` or calendar placeholder.
- Every VBA export has a matching unique `Attribute VB_Name` and `Option Explicit`; internal core modules also declare `Option Private Module`.
- Module inventory and public/core/UI/test naming boundaries match the approved architecture.
- The public-API manifest, MacroOptions manifest and source declarations contain the same 22 supported names/signatures.
- Every RibbonX callback resolves to a classified callback procedure.
- RibbonX and evidence schemas parse successfully.
- The fixture generator's `--check` mode reports no drift and the generated module is classified correctly.
- The documented test inventory includes `KPR_Test_Oracle.bas`.
- No generated Office binary, test-results output, `_Spill` API or calendar/business-day production placeholder is tracked.
- Existing repository checks continue to pass.

## Acceptance criteria

- [ ] Positive checks pass on the complete candidate tree.
- [ ] Focused negative self-tests fail for every new rule with actionable messages.
- [ ] Static checks do not claim that VBA imports, compiles or executes in Excel.
- [ ] The hosted workflow remains read-only, deterministic and pinned.
- [ ] The JSON artifact and Actions summary report all new rule results.
- [ ] Required-file policy does not require future `KPR_Cal*` modules.

## Dependencies

- [ ] #21
- [ ] #22
- [ ] #26

</details>

<details>
<summary><strong>#28 — Reconcile documentation, VERSION and CHANGELOG and assemble the candidate</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `documentation`, `blocked`, `P2`, `release`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/28

#### Body

## Objective

Commit the complete v0.0.2 release-candidate documentation and metadata before the exact-source Windows run.

## Scope

- Update README status, installation/compatibility guidance, contract links, public API and demo instructions.
- Document exactly 22 public names in one scalar/array-capable surface with no `_Spill` twins.
- State that scalar calls are supported on the Excel versions actually certified for scalar use, while multi-cell calls are supported and claimed only on dynamic-array Excel.
- Make no CSE compatibility claim until separately tested.
- Document the `1900-03-01 .. 9999-12-31` window, strict ISO parsing, optional-argument rule, 100,000-element cap and three native-error categories.
- Document the worksheet-Range date-system guard and the direct-VBA 1900 caller contract without any active-workbook fallback.
- State explicitly that calendars are v0.0.3 scope and business-day arithmetic/roll conventions are v0.0.4 scope.
- Correct the v0.0.1 changelog wording that describes a "scalar date and business-day source baseline"; v0.0.1 was repository setup only.
- Reconcile `CHANGELOG.md` with all post-v0.0.1 changes and prepare the v0.0.2 entry.
- Set `VERSION` to `0.0.2` only as part of the complete candidate.
- Document how to regenerate fixtures/demo and run certification.
- Keep generated Office binaries and final run output untracked.

## Acceptance criteria

- [ ] Documentation and manifests describe exactly the implemented 22-name surface.
- [ ] No page documents a spill twin or unsupported legacy multi-cell behavior.
- [ ] No page claims calendar or business-day support in v0.0.2.
- [ ] The v0.0.1 historical description is accurate without rewriting or moving its tag.
- [ ] `VERSION`, changelog and candidate scope agree.
- [ ] All repository and Markdown checks pass at the candidate commit.
- [ ] The candidate SHA is recorded for #29.
- [ ] No tracked change is planned after certification except a rerun-triggering correction.

## Dependencies

- [ ] #27

</details>

<details>
<summary><strong>#29 — Run exact-source Windows certification and publish v0.0.2</strong></summary>

- State: `open`
- Assignee: @danielep71
- Labels: `blocked`, `ci`, `P1`, `release`, `repository`, `tests`
- Milestone: `v0.0.2`
- URL: https://github.com/danielep71/KPR/issues/29

#### Body

## Objective

Certify the exact v0.0.2 candidate in real Windows Excel and publish only if every gate passes.

## Required certification

- Start from a clean checkout of the recorded candidate SHA.
- Import every production, test and demo VBE export into supported Windows Excel.
- Compile the VBA project with no compile error or missing reference.
- Run scalar cases through direct VBA calls under the documented 1900 contract.
- Exercise worksheet calls in a 1900 workbook.
- Exercise a 1904 workbook and prove that each public worksheet call returns one call-level `#N/A` rather than a shifted result.
- On dynamic-array Excel, run the complete multi-cell shape, broadcasting, optional-argument, per-element error and 100,000-element suites.
- Make no CSE or legacy multi-cell claim unless separately tested and evidenced.
- Run the full deterministic regression harness and cross-oracle suite.
- Run MacroOptions registration and best-effort unregistration twice.
- Load and exercise RibbonX.
- Install/use/remove the CommandBars UI twice.
- Generate the demo deterministically.
- Export the imported VBA source again and compare the normalized round trip.
- Record Excel/Windows environment, exact SHA, caller contexts, case totals and failures in the committed evidence schema.
- Keep run records under ignored `test-results/` and attach them to this issue and the GitHub pre-release.

## Exit gate

- [ ] Every issue in the v0.0.2 milestone is complete or this certification issue is the only remaining open item.
- [ ] Hosted static checks pass at the exact candidate SHA.
- [ ] Fresh import and VBA compilation pass in Windows Excel.
- [ ] All deterministic regression and allowed cross-oracle cases pass.
- [ ] Scalar/array parity, shapes, capacity, native errors, direct-VBA behavior and caller-state restoration pass.
- [ ] 1900 worksheet behavior and 1904 worksheet refusal pass.
- [ ] MacroOptions, Ribbon and CommandBars lifecycles are idempotent.
- [ ] Demo and source round-trip evidence are attached.
- [ ] Compatibility claims distinguish scalar support from dynamic-array-only multi-cell support and make no unevidenced CSE claim.
- [ ] Certification explicitly names calendars and business days as out of scope.
- [ ] No tracked file changes after the run; any change requires a new exact-source run.
- [ ] Create protected tag `v0.0.2` and a GitHub pre-release only after evidence is complete.
- [ ] Do not rewrite or move `v0.0.1`.
- [ ] Attach generated Office artifacts only as release assets; never commit them.

## Dependencies

- [ ] #28

</details>
