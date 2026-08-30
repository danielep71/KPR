# v0.0.2 implementation plan

Status: approved for issue-driven implementation. No production implementation or release certification has started.

- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)
- Planning baseline: `main` at `54d64ab3624aa051b19f9677aa49a3782cd26c60`
- Published protected baseline: `v0.0.1` at `abf38786eb48b3db1edced8ae26c756d9c7f5328`
- Version before candidate assembly: `0.0.1`

## Goal

Deliver the first functional KPR date layer as a source-first pre-release: Gregorian date primitives, strict parsing, scalar and dynamic-array APIs, registration and Excel UI infrastructure, deterministic tests, a reproducible demo builder, and exact-source Windows Excel certification.

The milestone is not complete merely because hosted static checks pass. It may be tagged and published only after the candidate source has been imported, compiled and executed successfully in Windows Excel and the evidence has been attached.

## Milestone sequence and boundaries

| Milestone | Scope |
| --- | --- |
| `v0.0.2` | Pure date primitives, parsing, scalar/spill APIs, MacroOptions, Ribbon, CommandBars, regression harness, demo and Windows certification |
| `v0.0.3` | Calendars: weekend masks, holiday sets and calendar composition |
| `v0.0.4` | Business-day arithmetic and roll conventions built on the calendar layer |

The `KPR_Cal_*` and `KPR_Cal_*_Spill` namespaces are reserved now. v0.0.2 creates no calendar placeholder modules and does not add weekend or holiday options to pure `KPR_Dates_*` functions.

## Supported API target

All supported functions are implemented in the two public facade modules. Members that must be technically `Public` for Excel, RibbonX, CommandBars or `Application.Run` remain infrastructure, not supported API.

### Scalar functions (22)

| Family | Functions |
| --- | --- |
| Inspection | `KPR_Dates_DayOfWeek`, `KPR_Dates_DaysInMonth`, `KPR_Dates_DaysInYear`, `KPR_Dates_IsMonthEnd`, `KPR_Dates_IsQuarterEnd`, `KPR_Dates_IsYearEnd`, `KPR_Dates_IsLeapYear` |
| Boundaries | `KPR_Dates_BeginOfMonth`, `KPR_Dates_EndOfMonth`, `KPR_Dates_BeginOfQuarter`, `KPR_Dates_EndOfQuarter`, `KPR_Dates_BeginOfYear`, `KPR_Dates_EndOfYear` |
| Arithmetic | `KPR_Dates_AddDays`, `KPR_Dates_AddWeeks`, `KPR_Dates_AddMonths`, `KPR_Dates_AddYears` |
| Weekday locators | `KPR_Dates_NthWeekdayOfMonth`, `KPR_Dates_LastWeekdayOfMonth` |
| Pillars | `KPR_Dates_PillarFromDates`, `KPR_Dates_DateFromPillar` |
| Host diagnostic | `KPR_Dates_HostDateSystem` |

`KPR_Dates_DateFromPillar` replaces the current plural `KPR_Dates_DatesFromPillar`; no compatibility alias is retained in this pre-release.

### Spill functions (19)

The following scalar members receive exact `_Spill` twins: `DayOfWeek`, `DaysInMonth`, `DaysInYear`, `BeginOfMonth`, `EndOfMonth`, `IsMonthEnd`, `IsQuarterEnd`, `IsYearEnd`, `IsLeapYear`, `AddDays`, `AddWeeks`, `AddMonths`, `AddYears`, `BeginOfQuarter`, `EndOfQuarter`, `BeginOfYear`, `EndOfYear`, `PillarFromDates`, and `DateFromPillar`.

`NthWeekdayOfMonth`, `LastWeekdayOfMonth` and `HostDateSystem` remain scalar-only in v0.0.2.

### Signature convention

- Required value arguments come first.
- Optional arguments use the `Opt_` prefix and come last in a stable order.
- Optional arguments are append-only within the owning domain.
- Inputs that may arrive from a worksheet use `Variant` so native Excel errors can be recognized and returned intentionally.
- Public results use `Variant` wherever a native Excel error is possible.
- The exact signatures and argument descriptions are frozen by issue #9 before behavior-changing implementation begins.

## Behavioral contract

### Inputs and errors

- String dates are locale-independent ISO `YYYY-MM-DD` only. Locale dates and numeric-looking strings are rejected.
- Numeric serials, VBA `Date`, integer arguments, Boolean, Empty, Null, Error, Range and in-memory array inputs follow one explicit type matrix.
- Incoming native Excel errors propagate verbatim.
- Invalid values, unsupported domains, shape mismatches and overflow return documented native Excel errors; they never return message strings or plausible-looking fallbacks.
- A blank required value is invalid and is not coerced to serial zero.

### Date systems and serials

- Worksheet calls identify the caller workbook rather than relying on an unrelated active workbook.
- Calls from a 1904-date-system worksheet are refused with `#N/A` to prevent a plausible 1,462-day shift.
- `KPR_Dates_HostDateSystem()` returns `1900` or `1904` for an identifiable worksheet caller and `#N/A` when the host cannot be established.
- Direct VBA numeric inputs use the documented 1900-system serial contract.
- The supported serial window and Excel serial-60 anomaly are frozen in the contract and applied before conversion to VBA `Date`.

### Arrays

- Scalar and 1×1 inputs expand to the target shape.
- Non-scalar arguments must have identical row and column dimensions.
- Row, column and rectangular orientation is preserved exactly.
- There is no row/column outer product or implicit cross-broadcasting.
- Worksheet Range orientation is preserved; the one-dimensional VBA-array rule is fixed by the contract.
- Invalid/error elements are handled independently, while a shape mismatch returns one intentional shape error.
- Every spill element delegates to the matching scalar function, making scalar/spill behavior a testable invariant.

The spill surface requires an Excel version with dynamic arrays. Legacy Excel may use the certified scalar surface; v0.0.2 does not promise legacy spill/CSE compatibility.

### Pillar rounding

`KPR_Dates_PillarFromDates` receives optional `Opt_Rounding`. Accepted values are case-insensitive `NEAREST` (default), `FLOOR` and `CEILING`.

- Absolute intervals shorter than seven days remain exact day pillars.
- Candidates are whole-week and calendar-month anchors built from the start date.
- `FLOOR` chooses the latest non-overshooting anchor.
- `CEILING` chooses the earliest non-undershooting anchor.
- `NEAREST` chooses the closest anchor; equivalent anchors prefer the month representation.
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
  KPR_Dates_Spill.bas
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
| `KPR_Core_Parse` | Strict scalar parsing and validation | `KPR_Core_Err` |
| `KPR_Core_Array` | Shape discovery, scalar expansion and elementwise traversal | `KPR_Core_Err` |
| `KPR_Dates` | Supported scalar facade | `KPR_Core_Err`, `KPR_Core_Parse`, `KPR_Core_Dates` |
| `KPR_Dates_Spill` | Supported spill facade | `KPR_Dates`, `KPR_Core_Array`, `KPR_Core_Err` |
| `KPR_Register` | MacroOptions manifest and repeatable registration | Public API names only; no date algorithms |
| `KPR_UI_Bars` | Temporary classic CommandBars build/teardown | Registration, demo and development test entry points |
| `KPR_UI_Ribbon` | Ribbon callbacks and invalidation | Registration, demo and development test entry points |

Core modules never depend on public facades, registration, UI, tests or demo code. Tests and the demo consume the supported API rather than private helpers.

## Registration and Excel UI

- `Application.MacroOptions` uses one category only: `KPR Dates`.
- One manifest covers all 41 supported functions, descriptions and argument descriptions.
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
  KPR_Test_Dates_Scalar.bas
  KPR_Test_Dates_Spill.bas
  KPR_Test_Dates_State.bas
test/fixtures/
  date-fixtures.tsv
test/evidence/
  README.md
  evidence-schema.json
  compile-template.md
tools/
  gen_fixtures.py
```

The canonical TSV and generated VBA fixture module come from an independent, deterministic Python generator with a drift-check mode. The runner is callable through `Application.Run`, returns Boolean success and writes a structured summary, case-level TSV and environment record for the supplied exact source SHA.

Coverage includes parsing/type rejection, serial limits, serial 60, leap years, all date boundaries, arithmetic clipping/preservation, pillar grammar and all rounding modes, both weekday bases, scalar/spill parity, all array shapes, blanks, mixed errors, overflow, repeatability, state restoration, host date-system diagnostics and 1904 refusal.

Cross-oracle checks are limited to `EOMONTH`, `EDATE`, `WEEKDAY`, `DAY`, `YEAR` and `MONTH` where the Excel and KPR contracts overlap. `WORKDAY.INTL` and `NETWORKDAYS.INTL` are out of scope.

Schemas and templates are tracked under `test/evidence/`. Actual runs are written under ignored `test-results/` and attached to the final certification issue and pre-release, avoiding an exact-SHA evidence recursion.

## Demo strategy

`demo/modules/KPR_Demo_Dates.bas` deterministically builds the demonstration workbook from a clean Excel instance and an explicit output path. The builder is authoritative source. Generated `.xlsx`, `.xlsm`, `.xlam` and other Office binaries are never committed; a certified demo or add-in may be attached later as a release asset.

The demo shows supported behavior, native errors and array shapes without claiming calendar support, business-day support, production readiness or unverified compatibility.

## Delivery phases

| Phase | Outcome | Issues |
| --- | --- | --- |
| 1. Contract and export baseline | Normative behavior and importable VBE source format | #9–#10 |
| 2. Core and scalar layer | Layered modules, strict parsing, date-system policy, pillar policy and 22 scalar functions | #11–#15 |
| 3. Array, spill and registration | Shared shape engine, 19 spill functions and MacroOptions manifest | #16–#18 |
| 4. Regression | Independent fixtures, runner, full suites and allowed Excel oracles | #19–#22 |
| 5. Demo and UI | Deterministic demo, RibbonX and classic CommandBars | #23–#25 |
| 6. Surface freeze and candidate | API classification, expanded static checks, docs, VERSION and CHANGELOG | #26–#28 |
| 7. Certification and release | Exact-source Windows import, compile, regression evidence and pre-release | #29 |

Parallel work is allowed only where the issue dependency list permits it. The `blocked` label remains on a dependent issue until its prerequisites are satisfied.

## Exit gate

- All v0.0.2 issues are complete or the certification issue is the sole remaining issue while evidence is assembled.
- The 41-name public surface, MacroOptions manifest, documentation and static inventory agree exactly.
- Hosted static checks pass at the exact candidate SHA without claiming Excel execution.
- A clean Windows Excel import and VBA compilation succeed at that same SHA.
- Full deterministic regressions, scalar/spill parity and allowed cross-oracles pass.
- 1900 behavior, 1904 refusal, UI lifecycles, demo generation, caller-state restoration and source round-trip are evidenced.
- Calendars and business-day calculations are named explicitly as out of scope.
- No tracked file changes occur after the exact-source run; any change requires a new run.
- Only then may `v0.0.2` be tagged and published as a GitHub pre-release.
- The protected `v0.0.1` tag is not rewritten or moved.

## Milestone issue register

This terminal register is generated from the live v0.0.2 milestone. It records every current milestone issue with its complete GitHub body as of this plan revision.

### [#9 — Specify the date-layer behavioural contract](https://github.com/danielep71/KPR/issues/9)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `documentation`, `behavior-change`, `P1`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Freeze the complete supported contract for the v0.0.2 date layer before production code is reorganized.

## Decisions to record

- Publish two distinct worksheet/VBA surfaces: 22 scalar functions under `KPR_Dates_*` and 19 elementwise spill functions using the `_Spill` suffix.
- Rename the existing plural `KPR_Dates_DatesFromPillar` member to `KPR_Dates_DateFromPillar`; do not retain a compatibility alias in this pre-release.
- Put required value arguments first and optional `Opt_` arguments last. Optional arguments are append-only within a domain.
- Reserve `KPR_Cal_*` and `KPR_Cal_*_Spill` for v0.0.3 and later. Do not add calendar options to the pure date functions.
- Accept locale-independent ISO `YYYY-MM-DD` text only; reject locale-formatted and numeric-looking strings.
- Define the accepted matrix for numeric serials, VBA `Date`, Boolean, Empty, Null, Error, text, Range values and in-memory arrays.
- Propagate incoming native Excel errors verbatim and return intentional native errors for all other invalid or unsupported inputs.
- Define the 1900/1904 policy, serial window, the Excel serial-60 anomaly, output types and overflow behavior.
- Define scalar expansion, exact-shape broadcasting, orientation preservation, blank handling, per-element errors and shape-mismatch behavior.
- Define the `NEAREST`, `FLOOR` and `CEILING` pillar policies and make `NEAREST` the default.
- Distinguish supported public API from procedures that must be technically `Public` for Excel, RibbonX, CommandBars, `Application.Run` or testing.

## Deliverable

Add a source-controlled date-layer contract document that is the normative reference for implementation, registration, tests and documentation.

## Acceptance criteria

- [ ] Every supported scalar and spill name has an exact VBA signature and return contract.
- [ ] The input-type and native-error matrices are explicit and contain no locale-dependent parsing path.
- [ ] Serial boundaries, serial 60, 1900/1904 behavior and direct-VBA behavior are unambiguous.
- [ ] Array orientation, scalar expansion, shape matching, blanks and per-element error propagation are specified.
- [ ] Pillar rounding is defined with tie-breaking, negative intervals and non-invariant round trips.
- [ ] The supported-API boundary and reserved `KPR_Cal_*` namespaces are recorded.
- [ ] Calendars, weekend masks, holidays, business-day arithmetic and roll conventions are explicitly out of scope.

## Dependencies

None. This issue is the normative dependency for all behavior-changing implementation work.

</details>

### [#10 — Adopt VBE export format and assert `Attribute VB_Name`](https://github.com/danielep71/KPR/issues/10)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `ci`, `P2`, `repository`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

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

### [#11 — Split `KPR_Dates_Days` into the layered architecture](https://github.com/danielep71/KPR/issues/11)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `blocked`, `code`, `P1`, `refactor`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Replace the monolithic `KPR_Dates_Days.bas` baseline with explicit core, public facade, registration and UI boundaries.

## Target production inventory

- `KPR_Core_Err.bas` — native error construction, classification and propagation helpers.
- `KPR_Core_Parse.bas` — strict date, integer and ISO parsing.
- `KPR_Core_Dates.bas` — date arithmetic, boundary logic and pillar core.
- `KPR_Core_Array.bas` — shape discovery, scalar expansion and elementwise application.
- `KPR_Dates.bas` — supported scalar date API only.
- `KPR_Dates_Spill.bas` — supported spill wrappers only.
- `KPR_Register.bas` — MacroOptions manifest and registration.
- `KPR_UI_Bars.bas` — classic CommandBars lifecycle.
- `KPR_UI_Ribbon.bas` — Ribbon callbacks only.

The registration and UI modules may be completed by their dedicated issues, but their ownership and dependency boundaries are fixed here.

## Dependency rule

`KPR_Dates_Spill` may call `KPR_Dates`, `KPR_Core_Array` and `KPR_Core_Err`. `KPR_Dates` may call `KPR_Core_Parse`, `KPR_Core_Dates` and `KPR_Core_Err`. Core modules must not depend on public facades, registration, UI, tests or demo code.

## Acceptance criteria

- [ ] Existing date behavior is migrated without leaving duplicate implementations.
- [ ] Supported functions exist only in `KPR_Dates` and `KPR_Dates_Spill`.
- [ ] Helpers are `Private` where possible and are never documented as supported API.
- [ ] Registration, callbacks, tests and demo builders do not contain financial/date algorithms.
- [ ] The old `KPR_Dates_Days.bas` file is removed after migration.
- [ ] No empty `KPR_Cal.bas` or `KPR_Core_Cal.bas` placeholder is created.
- [ ] The module ownership and allowed-dependency matrix is documented.

## Dependencies

- [ ] #9
- [ ] #10

</details>

### [#12 — Implement strict date and integer parsing with error propagation](https://github.com/danielep71/KPR/issues/12)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `behavior-change`, `blocked`, `code`, `P1`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Implement the input contract from #9 once, in the core parsing layer, so scalar and spill functions cannot diverge.

## Scope

- Parse ISO text only in exact `YYYY-MM-DD` form and validate components without DateSerial rollover.
- Reject locale-formatted dates, numeric-looking strings and permissive conversions.
- Parse integer arguments without silent truncation, Boolean coercion or overflow.
- Apply the approved behavior for numeric serials, VBA `Date`, Empty, Null, Boolean, Range values and in-memory values.
- Propagate incoming `CVErr` values unchanged.
- Return the contract's native Excel error for every rejected or unsupported value.
- Keep parsing free of UI, worksheet selection, active-workbook and locale state.

## Acceptance criteria

- [ ] No production path relies on locale-sensitive `CDate`, `IsDate` or equivalent permissive parsing.
- [ ] ISO leap-day and component validation is exact.
- [ ] Integer-required arguments reject fractions, Boolean values and overflow.
- [ ] Incoming `#N/A`, `#VALUE!`, `#NUM!` and other native errors propagate verbatim.
- [ ] Invalid input never yields a plausible date, a message string, a `MsgBox` or an unhandled VBA error.
- [ ] Deterministic unit fixtures cover every accepted and rejected input class.

## Dependencies

- [ ] #11

</details>

### [#13 — Enforce the 1900/1904 policy and add `HostDateSystem`](https://github.com/danielep71/KPR/issues/13)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `documentation`, `behavior-change`, `blocked`, `code`, `P1`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Make workbook date-system behavior explicit and prevent numerically plausible but shifted results.

## Contract

- Add scalar-only `KPR_Dates_HostDateSystem()`.
- Return numeric `1900` or `1904` when the worksheet caller's workbook can be identified; return `#N/A` when it cannot.
- Refuse worksheet UDF evaluation from a 1904-date-system workbook with `#N/A`.
- Treat direct VBA numeric inputs under the documented 1900-system serial contract.
- Apply the supported serial window and serial-60 policy defined in #9 consistently before converting to a VBA `Date`.
- Do not read or mutate unrelated active-workbook state as a substitute for identifying the caller.

## Acceptance criteria

- [ ] `HostDateSystem` is deterministic for 1900 and 1904 worksheet callers.
- [ ] Unknown/non-worksheet callers follow the documented `#N/A` path.
- [ ] Every scalar and spill worksheet UDF uses the same caller/date-system guard.
- [ ] 1904 worksheet calls cannot return 1,462-day-shifted values.
- [ ] Direct VBA behavior is documented and regression-tested.
- [ ] Serial boundaries, serial 60 and out-of-window values have native-error fixtures.
- [ ] Both workbook systems are exercised during Windows certification.

## Dependencies

- [ ] #12

</details>

### [#14 — Define and implement the three pillar-rounding modes](https://github.com/danielep71/KPR/issues/14)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `behavior-change`, `blocked`, `code`, `P2`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Replace the unresolved nearest-versus-floor behavior with one explicit pillar conversion policy shared by scalar and spill APIs.

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

- [ ] Scalar and spill paths call one pillar core.
- [ ] All three modes cover positive, negative, exact, tie and boundary cases.
- [ ] Month-end and leap-day anchors are deterministic.
- [ ] Invalid modes return the contract's native error.
- [ ] Formatting and parsing are mutually consistent for every supported pillar token.
- [ ] Regression fixtures explicitly demonstrate non-invariant rounded round trips.

## Dependencies

- [ ] #11
- [ ] #12

</details>

### [#15 — Complete the scalar date surface](https://github.com/danielep71/KPR/issues/15)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `enhancement`, `blocked`, `code`, `P2`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Deliver the complete supported scalar date API for v0.0.2 through `KPR_Dates.bas`.

## Supported scalar surface

Existing/re-signatured members: `DayOfWeek`, `DaysInMonth`, `DaysInYear`, `BeginOfMonth`, `EndOfMonth`, `IsMonthEnd`, `IsQuarterEnd`, `IsYearEnd`, `IsLeapYear`, `AddWeeks`, `AddMonths`, `AddYears`, `NthWeekdayOfMonth`, `LastWeekdayOfMonth`, `PillarFromDates`, and singular `DateFromPillar`.

New members: `AddDays`, `BeginOfQuarter`, `EndOfQuarter`, `BeginOfYear`, `EndOfYear`, and `HostDateSystem`.

## Acceptance criteria

- [ ] All 22 scalar functions use the exact signatures in #9 and return `Variant` where native errors are possible.
- [ ] Month, quarter and year boundaries are correct across leap years and year transitions.
- [ ] `AddDays`, `AddWeeks`, `AddMonths` and `AddYears` implement the documented clipping/preservation and overflow rules.
- [ ] Weekday bases and nth/last weekday locators reject unsupported arguments intentionally.
- [ ] Pillar functions use the policy from the pillar issue and expose only the singular `DateFromPillar` name.
- [ ] Shared core routines contain all substantive calculations; public functions are thin validation/facade layers.
- [ ] No calendar, holiday, weekend-mask, business-day or roll-convention function is introduced.

## Dependencies

- [ ] #12
- [ ] #13
- [ ] #14

</details>

### [#16 — Implement the array shape and broadcasting engine](https://github.com/danielep71/KPR/issues/16)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `blocked`, `code`, `P1`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Create one internal array engine with deterministic Excel orientation, broadcasting and per-element semantics.

## Contract

- Preserve row, column and rectangular two-dimensional shapes exactly.
- Treat a true scalar or 1×1 value as expandable to the target shape.
- Require all non-scalar operands to have identical row and column dimensions.
- Do not perform row-to-column outer products or implicit one-dimensional cross-broadcasting.
- Preserve worksheet Range orientation; normalize one-dimensional VBA arrays according to the rule frozen in #9.
- Reject multi-area, empty, jagged and unsupported array inputs intentionally.
- Treat a blank required value as invalid rather than coercing it to serial zero.
- Propagate native errors at the affected element without suppressing valid neighboring results.
- Return one intentional shape error when arguments cannot be broadcast.

## Acceptance criteria

- [ ] Scalar, row, column and rectangular fixtures preserve the documented output shape.
- [ ] Multi-argument scalar expansion works identically for ranges and in-memory arrays.
- [ ] Shape mismatches and unsupported array forms return the documented native error.
- [ ] Mixed-validity arrays preserve valid values and per-element errors.
- [ ] Array traversal is deterministic and does not select, activate or recalculate unrelated Excel state.
- [ ] Caller calculation, events, screen updating and selection state remain unchanged.
- [ ] Core array code contains no date-specific financial rules.

## Dependencies

- [ ] #11
- [ ] #12

</details>

### [#17 — Add the 19 spillable date UDFs](https://github.com/danielep71/KPR/issues/17)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `enhancement`, `blocked`, `code`, `P1`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Expose the elementwise dynamic-array date surface without changing the existing scalar names.

## Supported spill surface

Add `_Spill` twins for `DayOfWeek`, `DaysInMonth`, `DaysInYear`, `BeginOfMonth`, `EndOfMonth`, `IsMonthEnd`, `IsQuarterEnd`, `IsYearEnd`, `IsLeapYear`, `AddDays`, `AddWeeks`, `AddMonths`, `AddYears`, `BeginOfQuarter`, `EndOfQuarter`, `BeginOfYear`, `EndOfYear`, `PillarFromDates`, and `DateFromPillar`.

Do not add spill twins for `NthWeekdayOfMonth`, `LastWeekdayOfMonth` or `HostDateSystem` in v0.0.2.

## Acceptance criteria

- [ ] All 19 names use the exact `KPR_Dates_<Name>_Spill` convention.
- [ ] Every element delegates to the corresponding scalar public function and therefore shares parsing, date-system, calculation and native-error behavior.
- [ ] Scalar inputs, 1×1 inputs, rows, columns and rectangles follow the array-engine contract.
- [ ] Multi-argument UDFs support scalar expansion and reject incompatible shapes.
- [ ] Blank and error cells follow the per-element policy.
- [ ] Parity fixtures compare each spill element with a direct scalar call over the same inputs.
- [ ] Documentation states the dynamic-array Excel requirement and the legacy policy.

## Dependencies

- [ ] #15
- [ ] #16

</details>

### [#18 — Add the MacroOptions manifest and registration lifecycle](https://github.com/danielep71/KPR/issues/18)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `documentation`, `enhancement`, `blocked`, `code`, `P2`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Register the supported worksheet functions consistently without mixing descriptions and lifecycle code into calculation modules.

## Scope

- Maintain one manifest for the 41 supported functions, their category, descriptions and argument descriptions.
- Use exactly one MacroOptions category: `KPR Dates`.
- Add explicit, repeatable register-all entry points suitable for manual use, add-in startup and UI delegation.
- Keep infrastructure procedures technically public only where Excel requires it and classify them as unsupported infrastructure.
- Preserve caller/application state and avoid selections, activation, alerts or duplicate registration side effects.
- Keep calendar and business-day categories out of v0.0.2.

## Acceptance criteria

- [ ] Every supported scalar and spill UDF has exactly one manifest record.
- [ ] Function and argument descriptions match the normative signatures and optional-argument order.
- [ ] Repeated registration produces the same MacroOptions state and no duplicate UI artifacts.
- [ ] Registration can run through `Application.Run` without a worksheet caller.
- [ ] Failures return or record a controlled result and do not display message boxes during automated runs.
- [ ] Manifest completeness and uniqueness can be checked statically.
- [ ] Only `KPR Dates` is registered.

## Dependencies

- [ ] #15
- [ ] #17

</details>

### [#19 — Add the independent fixture generator and generated fixture module](https://github.com/danielep71/KPR/issues/19)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `blocked`, `ci`, `P1`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Generate deterministic expected results independently of the VBA implementation and make the fixture set reviewable in source control.

## Deliverables

- `tools/gen_fixtures.py`
- `test/fixtures/date-fixtures.tsv`
- `test/modules/KPR_Test_Fixtures_Generated.bas`

## Fixture contract

Each case must have a stable ID, suite, input-kind metadata, arguments, expected result type, expected value or native error, and a short rationale. The generated VBA module is derived from the canonical TSV and must not be hand-edited.

The generator may use Python standard-library date primitives, but must not invoke KPR code or mechanically translate the same VBA algorithm line for line.

## Acceptance criteria

- [ ] Generation is deterministic across repeated runs and platforms.
- [ ] A `--check` mode fails when committed TSV or VBA output is stale.
- [ ] Fixtures cover accepted/rejected input classes, serial edges, leap years, month/quarter/year boundaries, arithmetic edges, pillar modes and weekday bases.
- [ ] Expected native errors and output types are encoded explicitly.
- [ ] Fixture IDs and ordering remain stable unless a reviewed contract change requires an update.
- [ ] Generated VBA is a valid VBE export with `Option Explicit`.
- [ ] Calendar, weekend-mask, holiday and business-day fixtures are absent.

## Dependencies

- [ ] #12
- [ ] #14
- [ ] #15

</details>

### [#20 — Add the regression runner, assertions and evidence schema](https://github.com/danielep71/KPR/issues/20)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `blocked`, `code`, `P1`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Provide a deterministic, automation-friendly VBA test runner with structured evidence and complete caller-state restoration.

## Test modules

- `KPR_Test_Assert.bas`
- `KPR_Test_Runner.bas`
- committed evidence documentation and schema under `test/evidence/`

## Public test-infrastructure interface

- `KPR_Test_RunAll(ByVal SourceSha As String, ByVal OutputFolder As String) As Boolean`
- `KPR_Test_RunSuite(ByVal SuiteName As String, ByVal SourceSha As String, ByVal OutputFolder As String) As Boolean`

These entry points must be callable through `Application.Run`, but they are not supported production API.

## Evidence contract

Write a machine-readable summary, a case-level TSV and environment metadata covering the exact source SHA, Excel version/build/bitness, workbook date system, suite/case totals, failures and elapsed time. Keep schemas/templates under `test/evidence/`; write actual runs to ignored `test-results/` for attachment to the final certification issue/release.

## Acceptance criteria

- [ ] Assertions cover values, dates/serials, Booleans, strings, native error codes and array shape/content.
- [ ] A failed assertion records the case and continues unless the harness itself cannot proceed.
- [ ] Runs are deterministic apart from explicitly identified environment/time fields.
- [ ] The runner returns `True` only when every selected case passes.
- [ ] Calculation mode, events, screen updating, alerts, status bar, active workbook/sheet and selection are restored after success and failure.
- [ ] No automated path displays a `MsgBox` or leaves files/workbooks open.
- [ ] Evidence validates against the committed schema and records the supplied exact source SHA.

## Dependencies

- [ ] #19

</details>

### [#21 — Add scalar, parity, array, error and state regression suites](https://github.com/danielep71/KPR/issues/21)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `blocked`, `P1`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Exercise the complete supported date surface and prove scalar/spill parity without relying on visual inspection.

## Suites

- `KPR_Test_Dates_Scalar.bas`
- `KPR_Test_Dates_Spill.bas`
- `KPR_Test_Dates_State.bas`

## Required coverage

Accepted and rejected date inputs; serial boundaries and invalid serials; the serial-60 contract; leap years and 29 February; month, quarter and year boundaries; `AddDays`, `AddWeeks`, `AddMonths` and `AddYears`; clipping/preservation rules; pillar parse/format and all three rounding modes; non-invariant rounded round trips; weekday bases and weekday locators; all scalar-versus-spill pairs; scalar, 1×1, row, column and rectangular inputs; shape mismatches; blanks; mixed validity; native errors; overflow; repeatability; caller state; `HostDateSystem`; and 1904 refusal.

## Acceptance criteria

- [ ] Every supported scalar function has positive, edge and invalid-domain cases.
- [ ] Every one of the 19 spill functions is compared element-by-element with its scalar twin.
- [ ] Row, column and rectangle orientation is asserted, not inferred from displayed values.
- [ ] Mixed arrays prove that one invalid/error element does not corrupt valid neighbors.
- [ ] State restoration is tested after both passing and deliberately failing cases.
- [ ] Two consecutive runs produce identical case results.
- [ ] The suite contains no calendar, weekend-mask, holiday or business-day arithmetic case.

## Dependencies

- [ ] #17
- [ ] #20

</details>

### [#22 — Add Excel cross-oracle checks for the date surface](https://github.com/danielep71/KPR/issues/22)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `blocked`, `P2`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Compare KPR date primitives with independent native Excel worksheet functions where their contracts genuinely overlap.

## Oracle scope

Use only `EOMONTH`, `EDATE`, `WEEKDAY`, `DAY`, `YEAR` and `MONTH`. Construct comparisons only for inputs where Excel and KPR have matching documented semantics.

Do not treat Excel's permissive parsing, implicit coercion, serial-60 behavior or 1904 handling as authoritative when it differs from the KPR contract.

## Acceptance criteria

- [ ] Oracle cases state the exact overlap assumption being tested.
- [ ] Boundary and randomized deterministic samples cover month ends, leap years, arithmetic and weekday indices.
- [ ] A fixed seed and case order make repeated runs reproducible.
- [ ] Expected exclusions are documented rather than silently skipped.
- [ ] Oracle failures appear in the standard evidence output.
- [ ] `WORKDAY.INTL` and `NETWORKDAYS.INTL` are not used in v0.0.2.

## Dependencies

- [ ] #15
- [ ] #20

</details>

### [#23 — Add the deterministic date-demo builder](https://github.com/danielep71/KPR/issues/23)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `documentation`, `enhancement`, `blocked`, `code`, `P3`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Provide a reproducible demonstration workbook without committing an Office binary or making unsupported product claims.

## Deliverable

Add `demo/modules/KPR_Demo_Dates.bas` with a callable builder that creates a new workbook from a clean Excel instance, lays out fixed scalar and spill examples, applies deterministic formatting and saves only to an explicit output path.

The tracked builder is authoritative. A generated workbook may be attached later as a v0.0.2 release asset after exact-source certification.

## Acceptance criteria

- [ ] The builder creates the same sheets, labels, formulas, examples, widths, styles and named areas on repeated runs.
- [ ] Examples cover the supported date surface, array shapes, native errors and documented pillar behavior.
- [ ] The workbook labels calendars and business-day calculations as future scope, not missing supported features.
- [ ] No example claims certified accuracy, production readiness or legacy-Excel support beyond the agreed policy.
- [ ] The builder preserves Excel state and never overwrites an existing path without an explicit controlled error.
- [ ] No `.xlsm`, `.xlam`, `.xlsx` or other generated Office binary is committed.
- [ ] Demo generation is exercised during Windows certification.

## Dependencies

- [ ] #17
- [ ] #18

</details>

### [#24 — Add Ribbon callbacks, RibbonX and safe package injection](https://github.com/danielep71/KPR/issues/24)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `enhancement`, `blocked`, `code`, `P3`, `repository`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

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

### [#25 — Add idempotent classic CommandBars integration](https://github.com/danielep71/KPR/issues/25)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `enhancement`, `blocked`, `code`, `P3`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

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

### [#26 — Freeze the public-API manifest and classify infrastructure members](https://github.com/danielep71/KPR/issues/26)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `documentation`, `blocked`, `P2`, `repository`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Create one auditable inventory that separates supported worksheet/VBA API from internal and host-required public procedures.

## Deliverable

Add a machine-checkable public-API manifest and readable documentation covering:

- 22 supported scalar `KPR_Dates_*` functions;
- 19 supported `KPR_Dates_*_Spill` functions;
- private/internal `KPR_Core_*` helpers;
- MacroOptions registration entry points;
- Ribbon and CommandBars callbacks;
- regression-runner and demo-builder entry points.

Reserve `KPR_Cal_*` and `KPR_Cal_*_Spill` without creating placeholder modules or functions.

## Acceptance criteria

- [ ] Exactly 41 functions are classified as supported v0.0.2 API.
- [ ] Every technically public infrastructure member is explicitly marked unsupported/internal.
- [ ] Names, signatures and argument ordering agree with the contract and MacroOptions manifest.
- [ ] No supported function is implemented in a core, UI, registration, test or demo module.
- [ ] No unclassified public procedure remains.
- [ ] The manifest is suitable for static completeness and drift checks.
- [ ] Future calendar/business-day namespaces and milestone boundaries are recorded.

## Dependencies

- [ ] #24
- [ ] #25

</details>

### [#27 — Extend static checks for modules, manifests, fixtures and RibbonX](https://github.com/danielep71/KPR/issues/27)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `blocked`, `ci`, `P2`, `repository`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Extend the existing repository-integrity gate to cover the date-layer architecture while keeping Excel execution claims separate.

## New checks

- Required production, test, fixture, RibbonX, demo and contract files exist.
- Every VBA export has a matching unique `Attribute VB_Name` and `Option Explicit`.
- Module inventory and public/core/UI/test naming boundaries match the approved architecture.
- The public-API manifest, MacroOptions manifest and source declarations contain the same supported names/signatures.
- Every RibbonX callback resolves to a classified callback procedure.
- RibbonX and evidence schemas parse successfully.
- The fixture generator's `--check` mode reports no drift.
- No generated Office binary, test-results output or calendar/business-day production placeholder is tracked.
- Existing 11 repository checks continue to pass.

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

### [#28 — Reconcile documentation, VERSION and CHANGELOG and assemble the candidate](https://github.com/danielep71/KPR/issues/28)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `documentation`, `blocked`, `P2`, `release`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Commit the complete v0.0.2 release candidate documentation and metadata before the exact-source Windows run.

## Scope

- Update README status, installation/compatibility guidance, date contract links, public API and demo instructions.
- State explicitly that calendars are v0.0.3 scope and business-day arithmetic/roll conventions are v0.0.4 scope.
- Correct the v0.0.1 changelog wording that currently describes a "scalar date and business-day source baseline"; v0.0.1 was repository setup only.
- Reconcile `CHANGELOG.md` with all post-v0.0.1 changes and prepare the v0.0.2 entry.
- Set `VERSION` to `0.0.2` only as part of the complete candidate.
- Document the dynamic-array Excel requirement, legacy policy, ISO parsing, native errors, 1900/1904 behavior and unsupported surfaces.
- Document how to regenerate fixtures/demo and how to run certification.
- Keep generated Office binaries and final run output untracked.

## Acceptance criteria

- [ ] Documentation and manifests describe exactly the implemented 22 scalar and 19 spill functions.
- [ ] No page claims calendar or business-day support in v0.0.2.
- [ ] The v0.0.1 historical description is accurate without rewriting or moving its tag.
- [ ] `VERSION`, changelog and candidate scope agree.
- [ ] All repository and Markdown checks pass at the candidate commit.
- [ ] The candidate SHA is recorded for issue #29.
- [ ] No tracked change is planned after certification except a rerun-triggering correction.

## Dependencies

- [ ] #27

</details>

### [#29 — Run exact-source Windows certification and publish v0.0.2](https://github.com/danielep71/KPR/issues/29)

- State: `open`
- Assignee: [@danielep71](https://github.com/danielep71)
- Labels: `blocked`, `ci`, `P1`, `release`, `repository`, `tests`
- Milestone: [v0.0.2](https://github.com/danielep71/KPR/milestone/2)

<details>
<summary>Complete issue body</summary>

## Objective

Certify the exact v0.0.2 candidate in real Windows Excel and publish only if every gate passes.

## Required certification

- Start from a clean checkout of the recorded candidate SHA.
- Import every production, test and demo VBE export into supported Windows Excel.
- Compile the VBA project with no compile error or missing reference.
- Exercise 1900 and 1904 workbooks, including the deliberate 1904 UDF refusal.
- Run the full deterministic regression harness and cross-oracle suite.
- Run MacroOptions registration twice.
- Load and exercise RibbonX.
- Install/use/remove the CommandBars UI twice.
- Generate the demo deterministically.
- Export the imported VBA source again and compare the normalized round trip.
- Record Excel/Windows environment, exact SHA, case totals and failures in the committed evidence schema.
- Keep run records under ignored `test-results/` and attach them to this issue and the GitHub pre-release.

## Exit gate

- [ ] Every issue in the v0.0.2 milestone is complete or this certification issue is the only remaining open item.
- [ ] Hosted static checks pass at the exact candidate SHA.
- [ ] Fresh import and VBA compilation pass in Windows Excel.
- [ ] All deterministic regression and allowed cross-oracle cases pass.
- [ ] Scalar/spill parity, native errors, shapes, state restoration and repeatability pass.
- [ ] MacroOptions, Ribbon and CommandBars lifecycles are idempotent.
- [ ] Demo and source round-trip evidence are attached.
- [ ] The certification explicitly names calendars and business days as out of scope.
- [ ] No tracked file changes after the run; any change requires a new exact-source run.
- [ ] Create protected tag `v0.0.2` and a GitHub pre-release only after the evidence is complete.
- [ ] Do not rewrite or move `v0.0.1`.
- [ ] Attach generated Office artifacts only as release assets; never commit them.

## Dependencies

- [ ] #28

</details>
