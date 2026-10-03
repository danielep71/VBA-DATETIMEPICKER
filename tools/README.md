# DateTimePicker tooling

Requirements: Git, Python 3.10+ and Node.js 20+ on PATH. No Python packages or
Excel installation are needed for these checks. Run from the repository root:

```sh
python tools/check.py
```

This runs source/Ribbon/form integrity, procedure-scoped jump checks, conditional
declarations, whitespace, label fixtures and the tooling's negative tests.
Reports and logs go to ignored `test-results/`. Findings return a nonzero exit
code. CI uses `python tools/check.py --ci`, checking committed whitespace and
requiring a clean tree; PR runs also supply `--base <base-sha>`.
Both modes read tracked VBA from the working tree. Local results explicitly
record dirty state and cannot be used as clean committed release evidence.

The static workflow also runs checksum-pinned actionlint. CodeQL analyzes the
Python and JavaScript tooling, not VBA. PR analysis is read-only. None of these
checks compiles VBA, executes Excel or proves 32-bit runtime compatibility.

## Project-wide VBA call resolution (advisory)

```sh
python tools/check_vba_calls.py --root . --output test-results/check_vba_calls.json \
  --summary test-results/check_vba_calls.md
```

A diagnostic check, **not a release gate**: `check.py` does not run it, and the
static workflow runs it in a non-blocking step whose summary is added to the job
page. Its fixtures (`test_vba_calls.py`) do run inside `tool-tests`. It exists to
catch broken calls while the large modules are split (#24).

**What is analyzed.** `tools/vba-projects.json` declares each real VBA project
separately, never merged: `packaged-hosts` (both release packages),
`embedded-minimum` (the smallest compilable import in INSTALLATION.md) and
`production-core` (production source alone; non-gating, measuring the #86
boundary). Every configuration is analyzed in each environment of
`check_vba_conditionals.py`, so alternative `#If` declarations are never
combined. A finding present in only some environments is reported as
`configuration_dependent` and never as a definite defect.

**What it checks**, for targets it can resolve with certainty:

| Code | Finding |
| --- | --- |
| `VBA-CALL-001` | unqualified call to a project-named procedure that no module declares |
| `VBA-CALL-002` | `Module.Member` or a typed class receiver naming an undeclared member |
| `VBA-CALL-003` | Private member used from another module |
| `VBA-CALL-004` | unqualified name public in two standard modules (VBA "Ambiguous name") |
| `VBA-CALL-010`–`016` | too few/many arguments, omitted required, unknown, repeated or ParamArray named arguments, positional after named |
| `VBA-CALL-017` | property use needing an accessor (Get/Let/Set) that is not declared |
| `VBA-CALL-020` | conflicting declarations of one name in a module (Get/Let/Set of one property are one family, not a conflict) |
| `VBA-CALL-030` / `031` | literal `Application.Run`, `OnTime`, `OnKey`, `OnAction`, `CallByName`, `AddressOf`, macro-name or Ribbon target that is missing / only Private |
| `VBA-CALL-040` | advisory: same public name in more than one standard module |
| `VBA-CALL-090` / `091` | analysis failure / incomplete coverage |

Resolution follows VBA order: procedure locals and parameters, then the current
module, then public members of standard modules (`Option Private Module` keeps
them visible inside the project), then module names. An unresolved name is a
defect only when it matches `project_name_patterns` in the manifest; anything
else may be VBA, Excel, Office or MSForms and is counted as unknown.

**Severities.** `error` is a definite defect in a gating configuration; `warning`
is advisory or comes from a non-gating configuration; `unknown` could not be
established; `failure` means the analysis itself is incomplete. The status is
`fail` with any error, `incomplete` when a module failed to parse, a conditional
is indeterminate or coverage is incomplete, and `pass` only otherwise. An empty
list is never a clean verdict after a failure: call checks are skipped for that
configuration and the status stays `incomplete`.

**Deliberately unsupported**, counted as unknown: argument type compatibility;
receivers typed `Object`, `Variant` or a library type; chains after a member
whose type is not a project class; UserForm controls and built-in members;
unused-procedure detection (callbacks and event handlers are classified as
external entry points instead); callback and event signatures; document modules,
which are not exported; and project `#Const` symbols.

When a module is added, renamed or moved, update `tools/vba-projects.json` in the
same change; an unlisted or missing module makes the report `incomplete`.

## Inspect the VBA inside a built package (diagnostic)

```sh
python tools/inspect_vba_package.py --package <file.xlam|file.xlsm> \
  --source . --source-rev <tag-or-sha> \
  --output test-results/package-inspection.json --summary test-results/package-inspection.md
```

Written for the automatic-startup investigation (#113); **not a release gate**
and not run by `check.py` or CI (its fixtures run in `tool-tests`). It reads the
package without opening Excel or running a macro, and reports:

- **Identity:** package filename, size and SHA-256, the `xl/vbaProject.bin` hash,
  the resolved source SHA and the tool version and revision.
- **Inventory:** every embedded module, including `ThisWorkbook` and worksheet
  modules, with its kind and raw SHA-256.
- **Startup wiring:** `Workbook_Open`, `Workbook_AddinInstall`, `Auto_Open` and a
  Ribbon `onLoad` callback: placement, signature, and whether each reaches
  `DP_Start` directly or through project procedures, per `#If` environment.
  Paths through `Application.Run`/`OnTime`/`CallByName` or undeclared project
  names are unknown, and earlier calls under `On Error GoTo` are listed.
- **Source comparison** with the `--configuration` modules of
  `tools/vba-projects.json` (default `packaged-hosts`) at `--source-rev`.

Extraction uses the standard library only: the zip container, then the VBA
storage of `vbaProject.bin` ([MS-CFB], [MS-OVBA] decompression). It fails closed:
any structural surprise is `VBA-PKG-090`, and the inspection is then `failed`,
never a clean result. `--export-dir` accepts a folder of modules exported from the
VBE instead; that run is labelled `manual-export` (`VBA-PKG-092`) and stays
`incomplete`, because it does not prove what a package contains.

Comparison normalizes only: code page decoding, line endings, the export-only
header before `Attribute VB_Name`, trailing empty lines, and the `VB_Base`,
`VB_TemplateDerived` and `VB_Customizable` attributes that the VBE stores in
class and form streams but omits from exports. Raw hashes of both sides are kept.
A remaining difference is `VBA-PKG-012` unless it disappears once comments and
blank lines are removed (`VBA-PKG-015`).

The report keeps inspection completeness (`complete`, `incomplete`, `failed`)
separate from the findings outcome. `pass` needs a complete inspection and no
error. A hook that reaches `DP_Start` shows wiring, not that Excel calls it in a
given load mode: the compiled p-code, macro security, event state and add-in
load behaviour are not inspected. Only a fresh Excel session shows that.

| Code | Severity | Finding |
| --- | --- | --- |
| `VBA-PKG-001` | error | no correctly placed startup hook reaches `DP_Start` (unknown when a path is dynamic or analysis is incomplete) |
| `VBA-PKG-002` / `003` | error | lifecycle procedure misplaced / signature mismatch |
| `VBA-PKG-004` | warning | `DP_Start` reached only on a conditional path |
| `VBA-PKG-005` | unknown | dynamic or unresolved call on a startup path |
| `VBA-PKG-006` | info | hook text only in comments |
| `VBA-PKG-007` | warning | only `Workbook_AddinInstall` reaches `DP_Start` |
| `VBA-PKG-009` | info | an earlier call under `On Error GoTo` can skip `DP_Start` |
| `VBA-PKG-010` / `011` | error / warning | expected module missing / unexpected module |
| `VBA-PKG-012` / `015` | error / info | code differs / differs only in comments |
| `VBA-PKG-013` | info | document module with no tracked source |
| `VBA-PKG-014` | warning | module kind differs |
| `VBA-PKG-090` / `091` | failure | extraction failed / analysis incomplete |
| `VBA-PKG-092` | unknown | manual export, not package evidence |

## Prepare and check one release record

After building the two packages from the exact candidate, prepare one record:

```sh
python tools/check_release.py --prepare --candidate-sha <full-sha> \
  --assets-dir release-assets --output test-results/release-evidence.json
```

The command captures package names, sizes and hashes. Every Excel result starts
as `NOT RUN`; it never infers a pass. Existing evidence is never overwritten by
preparation. Fill this JSON with both hosts' compile results, environment,
standard and UI-smoke results, and the six manual matrix records. Set expected
test/suite totals from the frozen candidate's harness, not from a previous
release or the observed count alone. Record subcase coverage in `detail`; a
partially exercised matrix remains `NOT RUN`, with the passed subset described.

Environment fields are strings: Excel/Windows versions, bitness, scaling,
monitor count, load mode and other loaded add-ins. The expected package names
are `DATETIMEPICKER.vX.Y.Z.xlam` and `DATETIMEPICKER-demo-vX.Y.Z.xlsm`.

On a clean checkout of the candidate, run the static command in CI mode and:

```sh
git fetch origin main --tags
python tools/check.py --ci
python tools/check_release.py --evidence test-results/release-evidence.json \
  --assets-dir release-assets --static-report test-results/static-checks.json
```

After tagging, add `--require-tag` to verify the annotated tag and its target.
After publication, the **Verify published release evidence** manual workflow
accepts the candidate SHA and completed JSON, downloads the two published
packages and validates their hashes. It uses trusted `main` tooling and reads
the successful exact-candidate static workflow artifact (30-day retention),
without executing candidate code. It makes no repository changes. If that
artifact has expired, use the local commands with retained evidence; the hosted
workflow fails rather than substituting another commit's report.

The validator requires the candidate to be reachable from fetched `origin/main`
and records the main SHA used for that check. It rejects stale candidate IDs, missing hosts, failed compile,
incomplete/failed regressions, cleanup failures, changed bytes and unaccepted
manual gaps. `NOT RUN` needs a named `accepted_by` and a description; output is
`pass_with_limitations`, with the gaps retained individually. A reported `FAIL`
blocks validation. Acceptance remains a maintainer decision, not an automatic
waiver. Successful validation checks record consistency; it does not authenticate
the author, prove that the files were built from the claimed source, or verify
that a described manual test happened. Keep the original logs with the record.

Use this for future candidates containing these tools. Do not reconstruct or
re-certify v1.2.2: its accepted limitations, tag and packages remain authoritative.

## Template origin and adaptations

Adopted from `danielep71/EXCEL-VBA-PROJECT-TEMPLATE` at
`b903fe44ef6a032c4689870b83745afa1c22490d` under MIT (Daniele Penza):

- `_gatelib.py` and `check_committed_whitespace.py`: unchanged.
- `check_vba_conditionals.py`: Windows environments explicitly set `Mac=False`;
  the unknown-symbol fixture uses a genuinely unknown project symbol.
- `check_vba_jumps.py`: evaluates each reachable conditional source separately,
  retaining original line numbers for alternative procedure declarations.
- CodeQL workflow: unchanged. Action versions and actionlint pin follow the
  template; CodeQL action paths are grouped in Dependabot.

`check.py`, `check_source.py`, `check_release.py`, their tests and the static/release
workflows are DateTimePicker-specific adapters. They preserve `src/`, `demo/`
and `test/`; no initializer is run and no full template-contract adoption is claimed.
The canonical template's provisioning, portfolio, generated-project pilots and
Wiki publication machinery are not applicable to this existing repository.
