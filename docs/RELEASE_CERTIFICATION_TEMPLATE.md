# 🧾 Release Certification Record

**Certified by:** <!-- name -->
**Date:** <!-- YYYY-MM-DD -->

Use this human-readable record alongside `release-evidence.json`, not instead
of it. Generate the JSON with `tools/check_release.py --prepare` after testing
and hashing the final packages. Fill it from actual logs; never invent results.
An incomplete record is a blocker, not evidence of success.

## 1. Identity and order

| JSON field / identity | Recorded value |
| --- | --- |
| `schema_version` | `1` |
| `candidate_sha` | <!-- full 40-character main commit after integration --> |
| `version` | <!-- exact VERSION contents, X.Y.Z --> |
| `tag` | <!-- vX.Y.Z --> |
| Integration PR | <!-- link --> |
| Previous release tag | <!-- tag --> |
| Annotated tag object SHA | <!-- git rev-parse vX.Y.Z --> |
| Tag target SHA | <!-- git rev-parse vX.Y.Z^{commit}; must equal candidate_sha --> |

- [ ] Open the release branch at cycle start; collect reviewed task PRs.
- [ ] Freeze scope; finish version, dated changelog and documentation preparation.
- [ ] Integrate into main; record the resulting main commit as the candidate.
- [ ] Run static checks on the clean exact candidate; retain its successful main push run.
- [ ] Compile and exercise embedded candidate source in Excel.
- [ ] Build both packages from that SHA; include the required packaged harness.
- [ ] Reopen and compile each final package; run standard and UI-smoke packs.
- [ ] Complete the manual records; compute final sizes and hashes.
- [ ] Prepare and complete JSON; validate before tagging.
- [ ] Tag that exact certified commit; upload those same bytes without rebuilding.
- [ ] Verify published packages, tag and evidence with the release-check workflow.

The tag object and target are different identities. The target, source,
static report and both asset `candidate_sha` fields must all name the same
commit. Even a documentation-only new commit requires renewed candidate evidence.

## 2. Static evidence

Run `python tools/check.py --ci` on the clean candidate. Retain
`test-results/static-checks.json` and `.md`, complete output and workflow URL.
The JSON must have `status: pass`, exact `candidate_sha`, `dirty: false`,
`mode: committed`, and exactly these eleven checks, each with `exit_code: 0`:

| Required check | Result / evidence |
| --- | --- |
| `tool-tests` | <!-- --> |
| `check_vba_conditionals-fixtures` | <!-- --> |
| `check_vba_jumps-fixtures` | <!-- --> |
| `check_committed_whitespace-fixtures` | <!-- --> |
| `check_source` | <!-- --> |
| `check_vba_conditionals` | <!-- --> |
| `check_vba_jumps` | <!-- --> |
| `check_committed_whitespace` | <!-- --> |
| `check_markdown_links` | <!-- --> |
| `labels-sync` | <!-- --> |
| `labels-drift` | <!-- --> |

**Successful static-checks push run on main at candidate SHA:** <!-- URL -->

**Required separate Ribbon record:** run
`python .github/scripts/check-ribbon-resources.py --root src/ribbon` on the
same clean candidate and retain SHA, command, exit code and output here.
This is not an additional member of the eleven-check JSON set. A missing
path-filtered workflow run is not a pass; the local recorded result is required.
Missing Ribbon parts/resources block release.

```text
Candidate SHA:
Command / Python version / timestamp:
Exit code and output:
```

## 3. Assets and per-host environments

Both records are required in `assets`; store final files in `release-assets/`.

| Field | Add-in | Demo |
| --- | --- | --- |
| `host` | `xlam` | `xlsm` |
| `filename` | `DATETIMEPICKER.vX.Y.Z.xlam` | `DATETIMEPICKER-demo-vX.Y.Z.xlsm` |
| `candidate_sha` | <!-- same candidate --> | <!-- same candidate --> |
| `size` | <!-- integer bytes > 0 --> | <!-- integer bytes > 0 --> |
| `sha256` | <!-- final file digest --> | <!-- final file digest --> |
| `compile` | <!-- PASS required --> | <!-- PASS required --> |
| `environment.excel` | <!-- product/version/build --> | <!-- --> |
| `environment.windows` | <!-- edition/version/build --> | <!-- --> |
| `environment.bitness` | <!-- --> | <!-- --> |
| `environment.scaling` | <!-- --> | <!-- --> |
| `environment.monitors` | <!-- --> | <!-- --> |
| `environment.load_mode` | <!-- actual load/startup route --> | <!-- --> |
| `environment.other_addins` | <!-- explicit none or list --> | <!-- --> |

Record source import and package construction steps, including the `.frm/.frx`
pair, Ribbon XML, relationships and custom images. The packaged harness includes
`test/M_cDP_Test.bas` as specified in `tools/vba-projects.json`; its absence
cannot be replaced by a different smoke test. Remove no harness after testing.

## 4. Packaged regression evidence

Complete all four results. `standard` means `TST_DP_RunAll`; `ui_smoke` means
`TST_DP_RunAll_WithUISmoke` in an interactive desktop session. Expected totals
come from the candidate's suite topology, not by copying observed totals.

| JSON field | xlam standard | xlam ui_smoke | xlsm standard | xlsm ui_smoke |
| --- | --- | --- | --- | --- |
| `state` | <!-- PASS --> | <!-- PASS --> | <!-- PASS --> | <!-- PASS --> |
| `run` | <!-- integer --> | <!-- --> | <!-- --> | <!-- --> |
| `passed` | <!-- integer --> | <!-- --> | <!-- --> | <!-- --> |
| `failed` | <!-- 0 --> | <!-- 0 --> | <!-- 0 --> | <!-- 0 --> |
| `cleanup_failures` | <!-- 0 --> | <!-- 0 --> | <!-- 0 --> | <!-- 0 --> |
| `expected_run` | <!-- positive integer --> | <!-- --> | <!-- --> | <!-- --> |
| `suites` | <!-- positive integer --> | <!-- --> | <!-- --> | <!-- --> |
| `expected_suites` | <!-- positive integer --> | <!-- --> | <!-- --> | <!-- --> |

`run == passed == expected_run > 0` and `suites == expected_suites > 0` are
required for each pack. Paste all four runner summaries verbatim, plus logs,
timestamps and the source of expected counts. Compile/harness NOT RUN is blocking.

```text
xlam standard:
xlam ui_smoke:
xlsm standard:
xlsm ui_smoke:
Expected counts / topology evidence:
```

## 5. Manual records

The `manual` object must contain exactly these six names. Each needs `status`
(`PASS` or `NOT RUN`) and nonempty `detail`. For NOT RUN, also fill `accepted_by`
with the reviewer who explicitly accepted that gap. A failure is blocking;
never recast a failed test as NOT RUN. Accepted gaps yield `pass_with_limitations`.

| Name | Scenarios / evidence to describe | status | detail | accepted_by |
| --- | --- | --- | --- | --- |
| `provider-ownership` | second-provider refusal, owner teardown, clean-session DP_Show/DP_Close | <!-- --> | <!-- --> | <!-- --> |
| `clock-lifecycle` | real OnTime delivery, rename/requalification, close and restart | <!-- --> | <!-- --> | <!-- --> |
| `grid-ownership` | grid creation/removal, caller shapes preserved | <!-- --> | <!-- --> | <!-- --> |
| `application-state` | events/settings restored; write-back, tables, formulas, protection/array refusal | <!-- --> | <!-- --> | <!-- --> |
| `packaged-entry-paths` | keyboard, context menu, grid, Ribbon, startup; visuals/images and callbacks in both packages | <!-- --> | <!-- --> | <!-- --> |
| `ribbon-image-provenance` | tracked image sources and redistribution evidence, distinct from visual rendering | <!-- --> | <!-- --> | <!-- --> |

## 6. Validation and publication

```sh
python tools/check_release.py --prepare --candidate-sha <full-main-candidate-sha> --assets-dir release-assets --output test-results/release-evidence.json
# Fill the prepared record from the evidence above, then validate:
python tools/check_release.py --evidence test-results/release-evidence.json --assets-dir release-assets --static-report test-results/static-checks.json
# After the annotated tag exists:
python tools/check_release.py --evidence test-results/release-evidence.json --assets-dir release-assets --static-report test-results/static-checks.json --require-tag
```

| Record | Value |
| --- | --- |
| Validator output / status / limitations | <!-- --> |
| Published Release URL | <!-- --> |
| Published-package release-check workflow URL and result | <!-- --> |
| Downloaded filenames, sizes and SHA-256 comparison | <!-- --> |
| Installation / packaged entry-path verification | <!-- --> |
| Deviations and explicitly accepting reviewer | <!-- none or details --> |

The validator checks record consistency and file hashes; it does not execute
Excel, prove binary/source provenance or certify untested environments.
Retain source/build logs and disclose these boundaries in the release notes.

## 7. Sign-off

- [ ] All required JSON fields are filled from evidence; both hosts are present.
- [ ] The static report and separate Ribbon check describe the exact candidate.
- [ ] Every manual gap has an explicit acceptance and is disclosed.
- [ ] Tag target, tested sources, packages and evidence have one candidate identity.
- [ ] Published downloads match the tested bytes; no post-test rebuild occurred.
- [ ] Remaining limitations and deviations are recorded.

**Certified:** <!-- name, date; qualified by any disclosed limitations -->

## Related

- [Release procedure](../RELEASING.md)
- [Contribution rules](../CONTRIBUTING.md)
- [Installation](../INSTALLATION.md)
- [Tooling and validation limits](../tools/README.md)
