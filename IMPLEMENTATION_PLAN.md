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

## Milestone issues

| Issue | Title | Labels | Depends on |
| --- | --- | --- | --- |
| [#9](https://github.com/danielep71/KPR/issues/9) | Specify the date-layer behavioural contract | `documentation`, `behavior-change`, `P1` | — |
| [#10](https://github.com/danielep71/KPR/issues/10) | Adopt VBE export format and assert `Attribute VB_Name` | `repository`, `ci`, `P2` | — |
| [#11](https://github.com/danielep71/KPR/issues/11) | Split `KPR_Dates_Days` into the layered architecture | `refactor`, `code`, `P1`, `blocked` | #9, #10 |
| [#12](https://github.com/danielep71/KPR/issues/12) | Implement strict date and integer parsing with error propagation | `code`, `behavior-change`, `tests`, `P1`, `blocked` | #11 |
| [#13](https://github.com/danielep71/KPR/issues/13) | Enforce the 1900/1904 policy and add `HostDateSystem` | `code`, `documentation`, `behavior-change`, `tests`, `P1`, `blocked` | #12 |
| [#14](https://github.com/danielep71/KPR/issues/14) | Define and implement the three pillar-rounding modes | `code`, `behavior-change`, `tests`, `P2`, `blocked` | #11, #12 |
| [#15](https://github.com/danielep71/KPR/issues/15) | Complete the scalar date surface | `code`, `enhancement`, `tests`, `P2`, `blocked` | #12, #13, #14 |
| [#16](https://github.com/danielep71/KPR/issues/16) | Implement the array shape and broadcasting engine | `code`, `tests`, `P1`, `blocked` | #11, #12 |
| [#17](https://github.com/danielep71/KPR/issues/17) | Add the 19 spillable date UDFs | `code`, `enhancement`, `tests`, `P1`, `blocked` | #15, #16 |
| [#18](https://github.com/danielep71/KPR/issues/18) | Add the MacroOptions manifest and registration lifecycle | `code`, `enhancement`, `documentation`, `P2`, `blocked` | #15, #17 |
| [#19](https://github.com/danielep71/KPR/issues/19) | Add the independent fixture generator and generated fixture module | `tests`, `ci`, `P1`, `blocked` | #12, #14, #15 |
| [#20](https://github.com/danielep71/KPR/issues/20) | Add the regression runner, assertions and evidence schema | `tests`, `code`, `P1`, `blocked` | #19 |
| [#21](https://github.com/danielep71/KPR/issues/21) | Add scalar, parity, array, error and state regression suites | `tests`, `P1`, `blocked` | #17, #20 |
| [#22](https://github.com/danielep71/KPR/issues/22) | Add Excel cross-oracle checks for the date surface | `tests`, `P2`, `blocked` | #15, #20 |
| [#23](https://github.com/danielep71/KPR/issues/23) | Add the deterministic date-demo builder | `code`, `documentation`, `enhancement`, `P3`, `blocked` | #17, #18 |
| [#24](https://github.com/danielep71/KPR/issues/24) | Add Ribbon callbacks, RibbonX and safe package injection | `code`, `enhancement`, `repository`, `tests`, `P3`, `blocked` | #18, #20, #23 |
| [#25](https://github.com/danielep71/KPR/issues/25) | Add idempotent classic CommandBars integration | `code`, `enhancement`, `tests`, `P3`, `blocked` | #18, #20, #23 |
| [#26](https://github.com/danielep71/KPR/issues/26) | Freeze the public-API manifest and classify infrastructure members | `documentation`, `repository`, `P2`, `blocked` | #24, #25 |
| [#27](https://github.com/danielep71/KPR/issues/27) | Extend static checks for modules, manifests, fixtures and RibbonX | `ci`, `tests`, `repository`, `P2`, `blocked` | #21, #22, #26 |
| [#28](https://github.com/danielep71/KPR/issues/28) | Reconcile documentation, VERSION and CHANGELOG and assemble the candidate | `documentation`, `release`, `P2`, `blocked` | #27 |
| [#29](https://github.com/danielep71/KPR/issues/29) | Run exact-source Windows certification and publish v0.0.2 | `release`, `tests`, `ci`, `repository`, `P1`, `blocked` | #28 |
