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
