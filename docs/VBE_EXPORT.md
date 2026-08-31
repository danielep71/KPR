# VBE export format

Every tracked VBA file in this repository must conform to the repository's
Visual Basic Editor export format rather than being a hand-written
approximation of one. This page is the procedure contributors must follow so
that exported source, the repository static gate and a future Windows import
agree with each other.

This page describes a text format and a manual procedure. It makes no claim that
any tracked source imports, compiles or runs in Excel. That claim belongs to the
exact-source Windows certification recorded in issue
[#29](https://github.com/danielep71/KPR/issues/29).

## Format contract

A VBE export of a standard module begins with its component name:

```text
Attribute VB_Name = "KPR_DATES_DAYS"
```

The rules the static gate enforces:

| Rule | Requirement |
|---|---|
| Presence | Every `.bas`, `.cls` and `.frm` file declares `Attribute VB_Name`. |
| Form | The canonical VBE spelling `Attribute VB_Name = "<ComponentName>"`, with no leading whitespace and no alternative spacing. |
| Position | Line 1 for a `.bas` module. A `.cls` or `.frm` export opens with its own `VERSION`/`BEGIN` header block, so the declaration is accepted anywhere in the leading header region for those file types. |
| Identity | The declared name matches the file name exactly, including case. |
| Legality | The name is a legal VBA identifier of at most 31 characters. |
| Uniqueness | No two tracked VBA files declare names that are equal under VBA's case-insensitive component-name semantics. VBA components share one flat project namespace across `src/`, `test/` and `demo/`. |
| Declarations | `Option Explicit` is present, as it already was before this format was adopted. |

Procedure-level attributes such as `Attribute VB_Description` and
`Attribute VB_ProcData.VB_Invoke_Func` are rejected. Function and argument
descriptions are owned exclusively by the `Application.MacroOptions` manifest.
Those attributes are invisible in the editor but survive an export, so allowing
them would create a second, silent description mechanism that can disagree with
the manifest.

Encoding and line endings follow the policies already recorded in
`.gitattributes` and `.editorconfig`: exported VBA source is CRLF in the working
tree and ends with a newline. The VBE writes text in the Windows system code
page rather than UTF-8, so keep tracked VBA source ASCII-only. Typographic
quotes, en dashes and accented characters in comments are the usual way a file
stops round-tripping cleanly.

## Exporting from Excel

1. In the VBE, select the component in the Project Explorer.
2. Choose **File > Export File**.
3. Save into the directory that owns the component: `src/modules/` for
   production modules, `test/modules/` for regression source, `demo/modules/`
   for demonstration source.
4. Keep the file name identical to the component name. Do not rename the file
   afterwards and do not hand-edit the `Attribute VB_Name` value to something
   else; the two must agree, and the gate checks that they do.
5. Review the resulting diff before committing. An export that differs only in
   line endings or whitespace usually means the editor or a Git setting has
   overridden the repository policy.

## Importing into Excel

1. If a component of the same name is already loaded, **remove it first**:
   right-click the component, choose **Remove**, and decline the export prompt.
2. Choose **File > Import File** and select the exported file.

Importing a module while a component of the same name is already present does
not replace it. Excel keeps both and renames the newcomer, so `KPR_DATES_DAYS`
silently becomes `KPR_DATES_DAYS1` and the project then holds two copies of the
same procedures. Always remove before importing.

## Round trip

A normalized export/import round trip — importing the exact candidate source,
exporting it again and comparing the result — is release evidence, not a routine
contributor step. It is performed once against the recorded candidate SHA during
Windows certification and is deliberately out of scope for ordinary changes and
for the hosted static gate, which validates text shape only.
