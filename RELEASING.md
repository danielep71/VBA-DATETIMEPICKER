# 🚀 VBA DateTimePicker Release Guide

[![Release model: exact source](https://img.shields.io/badge/release-exact%20source-0969da)](#release-invariants)
[![Versioning: SemVer](https://img.shields.io/badge/versioning-SemVer-3f4551)](#versioning)
[![Evidence: required](https://img.shields.io/badge/evidence-required-success)](#evidence-record)
[![Security policy](https://img.shields.io/badge/security-policy-success)](SECURITY.md)
[![License: Apache 2.0](https://img.shields.io/badge/license-Apache_2.0-yellow.svg)](LICENSE)

This maintainer guide turns a reviewed commit into a traceable VBA DateTimePicker release. Source identity, validation, packaging, provenance, and publication must describe the same candidate.

> [!IMPORTANT]
> The release candidate must preserve **both** classes of package source:
> VBE/VBA inputs (including `UF_DatePicker.frm` with its adjacent
> `UF_DatePicker.frx` companion) and Open XML package inputs. The Ribbon is not
> created by VBE import or VBA compilation: `src/ribbon/customUI.xml`, its
> relationship part and every custom image it references must be injected into
> the Office package separately.

## 🧭 Release profile

| Property | Requirement |
| --- | --- |
| Project maturity | Versioned library with source and add-in deployment routes |
| Version source | [VERSION](VERSION) |
| Version scheme | Semantic Versioning |
| Tag format | Lower-case `vX.Y.Z` matching `VERSION` |
| Changelog | [CHANGELOG.md](CHANGELOG.md) |
| Installation contract | [INSTALLATION.md](INSTALLATION.md) |
| Vulnerability handling | [SECURITY.md](SECURITY.md) |
| License | [Apache 2.0](LICENSE) |

<a id="release-invariants"></a>
## 🔒 Release invariants

A release is valid only when all of these statements are true:

1. The candidate is identified by its full SHA and is reachable from protected `main`.
2. Release scope is frozen; unrelated work is deferred.
3. `VERSION`, changelog, package metadata, and current-version documentation agree.
   Module implementation-history headers are not release-version fields.
4. Static checks pass on the exact candidate.
5. VBA compiles and project-specific certification passes on that candidate.
6. Every binary artifact is built from that candidate, then tested as packaged.
7. Hashes and evidence bind artifacts to the candidate and tool environment.
8. The annotated lower-case tag points to the certified commit.
9. The GitHub Release uses that tag and the reviewed notes and assets.
10. Post-publication checks confirm that a new user can retrieve and validate it.

If an invariant is false, stop and repair the candidate. Never compensate by editing an already-tested artifact manually.

## 🗂️ Authority map

| Concern | Authoritative record |
| --- | --- |
| Current project version | `VERSION` |
| User-visible changes | `CHANGELOG.md` |
| Released source identity | Annotated Git tag and full commit SHA |
| Supported installation | `INSTALLATION.md` |
| Published binaries | GitHub Release assets |
| Integrity | SHA-256 hashes recorded with the release |
| Validation | Evidence record retained for the candidate |
| Vulnerability disclosure | `SECURITY.md` |

<a id="versioning"></a>
## 🧮 Versioning

Use `MAJOR.MINOR.PATCH` without a leading `v` inside `VERSION`.

- **MAJOR**: incompatible public API, behavior, file-format, or migration change.
- **MINOR**: backward-compatible capability.
- **PATCH**: backward-compatible correction or package/documentation fix included in the release.
- Use pre-release identifiers only when their support meaning is documented.

The Git tag adds the lower-case prefix: version `1.2.3` becomes annotated tag `v1.2.3`. Do not use upper-case `V`, moving tags, or a tag that differs from `VERSION`.

## ✅ Readiness review

- `M_DatePicker`, both classes, the UserForm pair, and the two `demo/` modules they need to compile ([#86](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/86)) match `INSTALLATION.md`.
- Only one provider implementation is active per Excel process.
- Every planned item is merged or explicitly deferred.
- Compatibility and migration consequences are understood.
- Security-sensitive work has completed private handling where necessary.
- The working tree, exported VBA sources, and non-VBE package inputs are reproducible.
- When the Ribbon is shipped, a clean checkout contains `src/ribbon/customUI.xml`
  plus every relationship part and custom image required to package it, and
  `.github/scripts/check-ribbon-resources.py` passes ([#90](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/90)).
- Maintainers and required reviewers are available.

Test both embedded-source and add-in consumption if both are advertised for the release.

<a id="one-command"></a>

## ⚡ One command: `tools/Release.cmd`

Sections 5 to 11 are run by one command on the Windows host
([#222](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/222)). Prepare
sections 1 to 4 as before: integrate into `main`, then check out the exact
integration commit with a clean tree. Run a dry run first:

```powershell
git fetch origin main --tags
git switch --detach origin/main
.\tools\Release.cmd -DryRun -Title "<short title>" -Refs 88,89 -EvidenceIssue <n>
.\tools\Release.cmd -Title "<short title>" -Refs 88,89 -EvidenceIssue <n>
```

It stops at the first failure and runs, in order:

1. **preflight**. The checkout must be exactly `origin/main` with a clean tree,
   and `VERSION` must have a dated CHANGELOG section. The tag and the GitHub
   Release must not exist yet, and the static-checks workflow must have a
   successful push run on `main` for this commit. GitHub CLI must be signed in,
   and Python with the documentation packages and Node.js must be available;
2. the static gates (`tools/check.py --ci`) and the Ribbon check, with its
   output kept (section 5);
3. both source runs, `tools/Certify-Excel.ps1` standard and `-WithUISmoke`
   (section 6);
4. the add-in and demo builds with the release file names; you answer their
   Ribbon questions (section 7);
5. both packs inside each package, `Certify-Excel.ps1 -Package` (section 7,
   step 10);
6. the user manual PDF, which you check page by page;
7. the six manual checks, asked one by one: PASS, NOT RUN with the name of whoever
   accepts the gap, or FAIL, which stops;
8. the evidence, with `tools/assemble_release.py`: `release-evidence.json`
   validated by `check_release.py`, the tag message and the release notes
   (sections 8 and 9);
9. a final confirmation: you type the tag;
10. the annotated tag, pushed, and the GitHub Release with the two packages, the
    PDF and `release-evidence.json` (sections 9 and 10). The published files
    are downloaded again and validated with `--require-tag`, and the
    **Verify published release evidence** workflow runs on its own (section 11).

`-DryRun` does steps 1 to 9 and stops before the tag; `-Rehearse` tries steps
2 to 7 on a checkout that is not a release candidate yet. Everything goes into one
folder next to the checkout, `DATETIMEPICKER-release-vX.Y.Z-<sha12>-<stamp>`,
ending with `release.txt` and `release.json`
([tools/README.md](tools/README.md#run-the-whole-release-windows)). The sections
below remain the reference for what each step checks and for doing a step by
hand.

## 1. Open the release branch, then freeze the candidate

### Start of the release cycle

Create `release/vX.Y.Z` from current protected `main` at the start of the cycle.
Collect reviewed task PRs there, following
[CONTRIBUTING.md](CONTRIBUTING.md#release-branches). Update every Dependabot
`target-branch` for the new cycle; it takes effect when integrated into `main`.

### Freeze and integrate

1. Stop feature work and review the complete delta from the previous release tag.
2. Complete version, changelog and documentation preparation (sections 2–4).
3. Pass PR checks and record final-head review, then integrate into `main` with
   a merge commit. Recreate the release branch at `main` if it was deleted;
   restore open task PR bases. No sync PR is needed after recreation at that SHA.
4. Fetch `origin/main`, check out the exact integration commit and record its
   full SHA as `candidate_sha`. This **main commit after integration** is the
   candidate to certify, build, test and tag; the pre-merge SHA is not interchangeable.
5. Run sections 5–7 on that clean candidate. Require a successful static-checks
   **push run on main** for this exact SHA, as `release-check.yml` requires.

```bash
git fetch --tags --prune
git rev-parse HEAD
git status --short
git diff --stat <previous-tag>..HEAD
git diff --name-status <previous-tag>..HEAD
```

Use **direct previous-tag → candidate endpoint semantics** for the release delta.
Do not use three-dot merge-base semantics here: release review asks what changed
between the source actually released last time and the source being proposed
now. Review the full diff as needed after the stat/name-status pass.

A dirty tree, unknown generated file, or unreviewed binary delta is blocking.

## 2. Synchronize version surfaces

Update every applicable surface in one reviewable change:

- `VERSION`
- `CHANGELOG.md`
- current-version package metadata (module headers carry no version; the demo
  builder header points to `VERSION`)
- README compatibility and installation examples
- Ribbon XML or package metadata when changed

Search for the prior version and retain only intentional historical references. Never rewrite historical changelog sections or immutable evidence.

## 3. Finalize the changelog

Move relevant entries from **Unreleased** into a dated `[MAJOR.MINOR.PATCH] - YYYY-MM-DD` section.

- Describe user-visible behavior, not commit mechanics.
- Use Added, Changed, Deprecated, Removed, Fixed, and Security where applicable.
- Link issues or pull requests when they improve traceability.
- Do not claim an artifact, platform, or guarantee that was not certified.
- Keep an empty Unreleased section for future work.
- Add or verify the comparison link for the new version.

## 4. Verify documentation and installation

- Follow [INSTALLATION.md](INSTALLATION.md) from a clean environment.
- Verify source paths, import order, module names, prerequisites, and upgrades.
- Confirm README examples use the supported API and current version.
- Check [CONTRIBUTING.md](CONTRIBUTING.md), [SECURITY.md](SECURITY.md), and license links.
- Remove stale commands and promises beyond the tested support matrix.

## 5. Run static gates

Run the local/CI tooling described in [tools/README.md](tools/README.md):

```sh
python tools/check.py --ci
```

Use one `release-evidence.json` for the candidate's package hashes, both-host
Excel results and disclosed manual gaps. `tools/check_release.py --prepare`
creates an untested record; validation never substitutes for Excel testing.
See the tooling guide for the preparation, validation and published-package
verification commands. This applies to future candidates containing the tools;
closed v1.2.2 evidence is not reopened.


- Validate exported VBA headers and file pairs.
- Treat VBE/VBA source and Open XML package source as separate completeness gates:

  ```text
  VBE/VBA source
    .bas / .cls / .frm + required .frx companions

  Open XML package source (src/ribbon/ mirrors the package's customUI/ folder)
    src/ribbon/customUI.xml               Office 2007 Ribbon format (2006/01)
    src/ribbon/_rels/customUI.xml.rels    image id -> package file mapping
    src/ribbon/images/*.png               every custom image the XML references
  ```

- When the Ribbon is shipped, run
  `python .github/scripts/check-ribbon-resources.py --root src/ribbon`. It proves that every
  custom `image` dependency has a tracked file and an explicit package mapping,
  and that the part's namespace matches its file name. The **Check Ribbon
  resources** workflow runs it on pull requests into `main` or `release/**` and
  on pushes to those branches when `src/ribbon/`, the script or the workflow
  changes. Missing Ribbon XML, relationships, or referenced
  custom resources are a **blocking package-source failure**.
  This is a required separate release record, not one of the eleven
  `static-checks.json` gates. Run it locally on the exact clean candidate and
  retain the command, SHA and successful output even when the path-filtered
  workflow has no run. Do not infer success from a skipped workflow.
- Do not treat **Debug → Compile VBAProject** as proof of Ribbon/package
  completeness. VBE compilation cannot see Open XML Ribbon parts.
- Check 32-bit and 64-bit declarations and scan for credentials or generated noise.

Capture commands, tool versions, timestamps, and complete results. Rerun affected gates after any change.

## 6. Certify in Excel

- Run **Debug → Compile VBAProject** on the exact candidate.
- Run `TST_DP_RunAll`.
  `tools/Certify-Excel.cmd` does both from a clean checkout in one step and
  writes an evidence folder naming the commit and host
  ([tools/README.md](tools/README.md#run-the-regression-pack-in-excel-windows)).
- Run `TST_DP_RunAll_WithUISmoke` in an interactive desktop session.
- Exercise `DP_Show` and `DP_Close`, including provider, write-back, and formula scenarios.

Certification rules:

- Use a clean workbook or documented clean fixture.
- Import only candidate files.
- Record Excel version, Windows version, and Office bitness.
- Test the advertised environment matrix.
- Treat warnings, repairs, or unexplained numerical deltas as failures.
- If the candidate SHA changes after certification, the prior exact-source
  evidence no longer certifies the new candidate; rerun the affected static and
  Excel gates.
- If the candidate SHA changes after release artifacts have been built, **discard
  the prior package-test and hash evidence**. Rebuild the artifacts from the new
  exact candidate, reopen and package-test them again, and compute new hashes
  only after those exact rebuilt files pass.

## 7. Build release artifacts

Planned outputs:

- `release-assets/DATETIMEPICKER.vX.Y.Z.xlam` add-in (required)
- `release-assets/DATETIMEPICKER-demo-vX.Y.Z.xlsm` demo (required)
- `release-assets/DATETIMEPICKER-user-manual-vX.Y.Z.pdf` user manual (required
  from `v1.2.3`)
- Source archive created by GitHub from the tag

For each artifact:

1. Start from a clean checkout/build location.
2. Use only candidate-controlled inputs.
3. Import the required VBE/VBA source, preserving `.frm`/`.frx` companions.
4. Inject required Open XML package parts separately. For the Ribbon, copy
   `src/ribbon/` into the package's `customUI/` folder unchanged:

   ```text
   src/ribbon/customUI.xml                  -> customUI/customUI.xml
   src/ribbon/_rels/customUI.xml.rels       -> customUI/_rels/customUI.xml.rels
   src/ribbon/images/DP_GridIcon_64.png     -> customUI/images/DP_GridIcon_64.png
   src/ribbon/images/reset.png              -> customUI/images/reset.png
   src/ribbon/images/demo.png               -> customUI/images/demo.png
   ```

   The package's root relationship must use the Office 2007 extensibility type
   (`http://schemas.microsoft.com/office/2006/relationships/ui/extensibility`)
   for `customUI/customUI.xml`. Office RibbonX Editor does this when the part is
   added as an "Office 2007 Custom UI Part".

   For the add-in, `tools/Build-Addin.cmd` performs steps 1 to 6 and 8, the
   Ribbon part of step 9 and step 11 from a clean checkout. The add-in's
   `ThisWorkbook` code is a candidate input too. It is tracked as
   `src/addin/ThisWorkbook.cls`, and the tool copies it into the workbook's
   document module, because a document module cannot be imported. The tool
   records its evidence in a folder outside the checkout
   ([tools/README.md](tools/README.md#build-the-add-in-and-check-its-ribbon-windows)).
   `tools/Certify-Excel.cmd -Package <file>` performs step 10 on either package
   ([tools/README.md](tools/README.md#run-the-regression-pack-inside-a-built-package-windows)).
   `tools/Build-Demo.cmd` builds the demo workbook the same way, with the demo
   sheet very hidden behind a Start sheet
   ([tools/README.md](tools/README.md#build-the-demo-workbook)).
5. Fail the build if a required Ribbon part or referenced custom resource cannot
   be resolved from the candidate-controlled inputs.
6. Compile the VBA project before saving; remember that compilation validates
   VBA, **not** Open XML package completeness.
7. Include the packaged harness required by `tools/vba-projects.json`, including
   `test/M_cDP_Test.bas`, in both packages. Removing it or substituting another
   smoke test breaks the certification rule. `check_release.py` cannot detect
   it: the validator checks package names, sizes, hashes, build-machine
   metadata and the recorded results, not the VBA inside the package.
8. Remove the build machine's metadata: the save folder, author, sensitivity
   label, document-library and printer data that Excel records on save
   ([#221](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/221)).
   Run `python tools/sanitize_package.py <package> --in-place` on each package
   before it is reopened, tested and hashed. `tools/Build-Addin.cmd` does it for
   the add-in, and `check_release.py` refuses a package that still carries any
   of it.
9. Reopen the packaged artifact and verify the shipped Ribbon itself: expected
   group/controls are present, all intended custom images render, and each
   enabled Ribbon action dispatches correctly.
10. Run both `TST_DP_RunAll` and `TST_DP_RunAll_WithUISmoke` on each package.
    Record expected and actual run/suite totals, failures and cleanup failures.
    The harness writes the suites dispatched and completed to `K9:K10` of
    `TST_DP_RESULTS`. `tools/Certify-Excel.cmd -Package <file>` runs both packs
    in the package without saving it and fails if its SHA-256 changes.
11. Record filename, size, and SHA-256.
12. Never edit the artifact after hashing.

The user manual PDF is rendered from `docs/USER_MANUAL.md` at the exact
candidate. Its links to other repository files point at that commit. It has
page numbers, bookmarks, and a table of contents with the page of each section.
Install Python-Markdown and pypdf from the hash-locked
`tools/requirements-docs.txt` first, then render with headless Chrome or Edge:

```sh
python -m pip install --require-hashes --only-binary=:all: -r tools/requirements-docs.txt
python tools/render_manual.py --html test-results/USER_MANUAL.html --pdf release-assets/DATETIMEPICKER-user-manual-vX.Y.Z.pdf
```

Open the PDF and check every page before you hash it. Record its filename, size
and SHA-256 in the certification record (section 3 of the template).
`check_release.py` validates the two packages only, not the manual.

The Ribbon source above was recovered from the published `v1.2.2` add-in
([#90](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/90)). Before that, the repository tracked an Office 2010 `customUI14.xml`
that no release had shipped, and none of the images or the relationship part.
This guide must not encode private/local resource locations as a workaround.

```powershell
Get-FileHash -Algorithm SHA256 .\release-assets\<artifact>
```

```bash
sha256sum release-assets/<artifact>
```

Publish only artifacts promised by the installation guide.

## 8. Validate the evidence before tagging

Integration already happened in section 1. Do not merge or edit the candidate
between final certification and tagging. Prepare the machine-readable record
from the exact tested package bytes, then fill it from the retained logs:

```sh
python tools/check_release.py --prepare --candidate-sha <full-main-candidate-sha> --assets-dir release-assets --output test-results/release-evidence.json
python tools/check_release.py --evidence test-results/release-evidence.json --assets-dir release-assets --static-report test-results/static-checks.json
```

`tools/assemble_release.py` fills the record from the evidence folders instead
([tools/README.md](tools/README.md#assemble-the-record-from-the-evidence-folders));
`tools/Release.cmd` runs it (step 8 of [One command](#one-command)).

Use [the certification template](docs/RELEASE_CERTIFICATION_TEMPLATE.md) alongside
that JSON. Record both per-host environments, standard/UI-smoke counts and all
six manual records. Blank or NOT RUN compile/harness fields block validation.
A manual NOT RUN requires a nonempty reason and named accepting reviewer; it
produces `pass_with_limitations`, not an unqualified pass. Unresolved failures
cannot be accepted as gaps. The validator checks consistency and local hashes,
not whether Excel was actually run or the binary was built from the asserted SHA.

A new commit, including documentation changes, requires a new candidate and
renewed exact-source evidence; package changes invalidate test and hash evidence.

## 9. Create the annotated tag

Tag only the certified commit on `main`.

The tag message is the release's permanent, offline record: it travels with every
clone, while Release pages and issues can be edited later. `tools/assemble_release.py`
writes it from the evidence record as `tag-message.txt`, and `tools/Release.cmd`
creates and pushes the tag from it. By hand, write it into a file outside the
repository, using this template:

```text
vX.Y.Z - <short release title>

Certified at <full 40-character SHA>.
Embedded .xlsm: standard <passed>/<run>, UI smoke <passed>/<run>, Failed=0, CleanupFailures=0.
Packaged .xlam: standard <passed>/<run>, UI smoke <passed>/<run>, Failed=0, CleanupFailures=0.
Refs #<n>, #<n>.
Deviations: none.
Evidence: #<certification issue>.
```

- Use the plain hyphen in the first line, as `v1.2.1` does.
- Copy the totals verbatim from the runner summary lines; never round, infer or
  reuse a previous release's figures.
- `Deviations:` is mandatory. Write `none`, or list each gate that was not run
  (for example `#50 manual matrix not run`). A deviation accepted after tagging
  belongs in the certification issue and the changelog, because the tag cannot
  be changed.
- Keep the file free of credentials, local paths and machine names.

```bash
git switch main
git fetch origin main --tags
git rev-parse <certified-main-sha>
git tag -a vX.Y.Z <certified-main-sha> -F ../tag-message-vX.Y.Z.txt
git show --no-patch --decorate vX.Y.Z
git push origin vX.Y.Z
```

Create the tag from the command line. Creating it from the GitHub Release form
produces a lightweight tag with no message and no tagger.

Before pushing, confirm the tag equals `VERSION`, targets the certified commit,
is annotated (`git cat-file -t vX.Y.Z` prints `tag`), shows the intended message,
and has a matching dated changelog section. Never delete and recreate a public
tag to hide a mistake.

<a id="evidence-record"></a>
## 🧾 Evidence record

Retain at least:

- candidate commit SHA
- `release-evidence.json` and the validator output
- exact-candidate `static-checks.json` with all eleven gates and successful main push-run URL
- separate Ribbon-resource command, candidate SHA and successful output
- all six named manual records: `provider-ownership`, `clock-lifecycle`,
  `grid-ownership`, `application-state`, `packaged-entry-paths`,
  `ribbon-image-provenance`; status, detail and accepting reviewer for every gap
- all seven environment fields for each package host: `excel`, `windows`,
  `bitness`, `scaling`, `monitors`, `load_mode`, `other_addins`
- supported Excel versions and bitness
- core and UI-smoke results
- artifact SHA-256 hashes
- version and tag
- previous release tag
- release-preparation pull request
- artifact filenames and sizes
- validation timestamps
- known deviations and approving reviewer

A release note summarizes evidence; it does not replace it.

## 10. Publish the GitHub Release

Create the release from the annotated tag and include:

1. User-facing summary and highlights.
2. Upgrade or migration notes.
3. Supported platform statement.
4. Known limitations.
5. Installation link.
6. Asset table with SHA-256 hashes.
7. Full changelog comparison link.
8. Security-reporting link.

Upload the already-hashed artifacts and `release-evidence.json`. Do not rebuild
between tagging and upload. `tools/assemble_release.py` writes the notes
(`release-notes.md`): the dated CHANGELOG section, the downloads with SHA-256,
both packs on both packages, the environment, the manual checks and the
deviations. Add any upgrade notes or highlights by editing the CHANGELOG section
before the candidate is frozen, not the generated file.

Publishing the release starts `release-check.yml` on its own. It takes the
candidate from the tag and the record from the `release-evidence.json` asset,
retrieves the successful main push-run static report and both published
packages, then validates the annotated tag and downloaded hashes. It can also be
dispatched by hand with `candidate_sha` and the completed JSON. Retain its
report and disclose any accepted manual limitations.

## 11. Verify after publication

| Check | Result |
| --- | :---: |
| Tag resolves to the certified SHA | ☐ |
| `VERSION` and changelog match the tag | ☐ |
| Assets download and hashes match | ☐ |
| Installation links and examples work | ☐ |
| Packaged artifact opens and passes its smoke test | ☐ |
| Packaged Ribbon is present; intended custom images render and callbacks dispatch | ☐ |
| Source archive contains expected VBA **and Open XML/Ribbon resource inputs** | ☐ |
| Default branch is ready for the next Unreleased cycle | ☐ |

Do not announce broad availability until these checks pass.

## 🧯 Recovery

### Before tag publication

Fix the branch or pull request, update evidence, and rerun affected gates. Rebuilding invalidates the prior hash and packaged-test result.

### After tag, before GitHub Release

Pause publication and document the correction. Prefer a new patch candidate whenever the tag may have propagated.

### After public release

Do not silently replace assets or move the tag.

- Mark a dangerous release clearly and remove assets only for user safety.
- Publish a corrected patch with a new tag.
- Record the problem and remedy in the changelog.
- Use [SECURITY.md](SECURITY.md) for vulnerabilities.
- Preserve evidence explaining what users received.

## ☑️ Final maintainer checklist

| # | Gate | Done |
| ---: | --- | :---: |
| 1 | Scope frozen, diff reviewed and preparation integrated into main | ☐ |
| 2 | Version surfaces synchronized | ☐ |
| 3 | Changelog finalized | ☐ |
| 4 | Installation verified cleanly | ☐ |
| 5 | Static checks pass | ☐ |
| 6 | Excel certification passes | ☐ |
| 7 | Artifacts built and packaged-tested | ☐ |
| 8 | Hashes and evidence recorded | ☐ |
| 9 | Evidence validated against the exact certified main SHA | ☐ |
| 10 | Annotated tag targets certified main SHA | ☐ |
| 11 | GitHub Release published | ☐ |
| 12 | Post-publication checks pass | ☐ |

## 📚 Related documents

- [README.md](README.md) — project overview
- [INSTALLATION.md](INSTALLATION.md) — deployment and removal
- [CONTRIBUTING.md](CONTRIBUTING.md) — change and review workflow
- [CHANGELOG.md](CHANGELOG.md) — release history
- [SECURITY.md](SECURITY.md) — supported versions and private reporting
- [VERSION](VERSION) — authoritative version
- [LICENSE](LICENSE) — Apache 2.0 license

---

**Release principle:** certify one exact source revision, build from it once, and publish only evidence-backed artifacts derived from it.
