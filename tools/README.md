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
