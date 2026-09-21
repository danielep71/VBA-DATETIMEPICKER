# 🧾 Release Certification Record

**Version:** `vX.Y.Z`
**Certified by:** <!-- name -->
**Date:** <!-- YYYY-MM-DD -->

Fill this in as you certify, not afterwards from memory. Every field is a fact
someone else must be able to check without asking you. A blank is better than a
guess — a blank is visible, a guess is not.

Procedure and rules live in
[RELEASING.md](https://github.com/danielep71/VBA-DATETIMEPICKER/blob/main/RELEASING.md);
this file records what that procedure produced.

---

## 1. Identity

| Field | Value |
| --- | --- |
| Reviewed source SHA | <!-- full 40-character SHA of the source the evidence describes --> |
| Tag name | <!-- e.g. vX.Y.Z --> |
| Annotated tag object SHA | <!-- `git rev-parse vX.Y.Z` --> |
| Tag target SHA | <!-- `git rev-parse vX.Y.Z^{commit}` --> |
| Previous release tag | <!-- e.g. vX.Y.Z-1 --> |
| `VERSION` file contents | <!-- must equal the version above --> |
| Branch | <!-- e.g. release/vX.Y.Z --> |

> [!IMPORTANT]
> An annotated tag is its own git object. `git rev-parse vX.Y.Z` returns the tag
> object; `git rev-parse vX.Y.Z^{commit}` returns the commit it targets. Both are
> recorded because only the target can be compared against the reviewed source,
> and quoting the tag object SHA as though it were the commit is wrong in a way
> nothing surfaces.
>
> The reviewed source SHA and the tag target SHA are also recorded separately.
> They are usually the same. When they are not — because documentation-only
> commits landed after the last executable change — say which commit the
> regression figures actually describe. Figures quoted against a commit that did
> not produce them look like evidence and are not.

If the reviewed source SHA and the tag target SHA differ, state why:

```text
<!-- e.g. documentation-only commits between the reviewed source and the tag -->
```

---

## 2. Environment

Record only what you actually ran on.

```text
Excel product, version, build:
Office bitness:                    32-bit / 64-bit
Windows edition, version, build:
Display scaling / monitors:
Other add-ins loaded in the process:
Clean VM or developer workstation:
```

> [!NOTE]
> A developer workstation with other add-ins loaded is an acceptable environment
> if it is recorded as one. What is not acceptable is leaving the reader to
> assume a clean machine.

---

## 3. Required order

Tick these in sequence as you go. The order is the guarantee: certifying a
candidate that was built before it was frozen proves nothing about the artifacts
that ship, and hashing before the package is tested means the hash may belong to
a file that fails.

- [ ] 1. Freeze the candidate and record its SHA.
- [ ] 2. Static gates pass on that exact SHA.
- [ ] 3. Certify in Excel on that exact SHA — embedded host.
- [ ] 4. Build the release artifacts **from that same SHA**.
- [ ] 5. Package-test the built artifacts — packaged host.
- [ ] 6. Hash the tested artifacts, not earlier files.
- [ ] 7. Tag the certified commit.
- [ ] 8. Upload the already-hashed artifacts. Do not rebuild.
- [ ] 9. Re-read the published release and verify it.

If any step sends you backwards, every later step is void and must be redone.
Record what was redone in **Deviations**.

---

## 4. Static gates

| Gate | Result |
| --- | --- |
| `Debug → Compile VBAProject` | <!-- PASS / FAIL --> |
| Encoding: source files are ASCII | <!-- PASS / FAIL --> |
| Line endings match `.gitattributes` | <!-- PASS / FAIL --> |
| No test-only or temporary instrumentation in the candidate | <!-- PASS / FAIL --> |
| VBE/VBA source companions complete (`.frm` + required `.frx`) | <!-- PASS / FAIL --> |
| Ribbon/Open XML package-input inventory complete | <!-- PASS / FAIL / N/A --> |
| Every RibbonX custom image/resource resolves to tracked candidate source | <!-- PASS / FAIL / N/A --> |
| Working tree clean at the reviewed SHA | <!-- PASS / FAIL --> |

Commands run and their output:

```text
<!-- e.g. git status --short ; git grep for probe markers -->
```

---

## 5. Regression evidence — embedded `.xlsm`

Paste the runner's own summary line, verbatim. Do not retype the counts into
prose: the line is the evidence, a transcription is a claim about it.

```text
Host workbook:     <!-- filename, IsAddin=False -->

TST_DP_RunAll
<!-- INFO | Harness | Summary | State=PASS; Run=###; Passed=###; Failed=0; CleanupFailures=0 -->

TST_DP_RunAll_WithUISmoke
<!-- INFO | Harness | Summary | State=PASS; Run=###; Passed=###; Failed=0; CleanupFailures=0 -->

Suite topology:    <!-- ## standard / ## with UI smoke -->
```

---

## 6. Regression evidence — packaged `.xlam`

**This section is required.** It has no default and is not satisfied by the
embedded results above.

```text
Package filename:  <!-- exact name as it will be published -->

TST_DP_RunAll
<!-- INFO | Harness | Summary | State=PASS; Run=###; Passed=###; Failed=0; CleanupFailures=0 -->

TST_DP_RunAll_WithUISmoke
<!-- INFO | Harness | Summary | State=PASS; Run=###; Passed=###; Failed=0; CleanupFailures=0 -->

Suite topology:    <!-- ## standard / ## with UI smoke -->
```

### Packaged Ribbon / Open XML verification

VBE compilation is not evidence that RibbonX was packaged. Record the actual
packaged result:

| Check | Result |
| --- | --- |
| Expected Ribbon group/controls are present | <!-- PASS / FAIL / N/A --> |
| `customUI14.xml` is present in the Office package when advertised | <!-- PASS / FAIL / N/A --> |
| Required Ribbon relationships/package metadata resolve | <!-- PASS / FAIL / N/A --> |
| Every intended custom Ribbon image renders | <!-- PASS / FAIL / N/A --> |
| Each enabled Ribbon callback dispatches correctly | <!-- PASS / FAIL / N/A --> |

Candidate-controlled Ribbon inputs used:

```text
<!-- paths to customUI14.xml, relationships/mapping and custom image resources -->
```

If the harness is not present in the package, say so here and record how the
package was exercised instead:

```text
<!-- N/A if the harness ran; otherwise describe what was run -->
```

> [!IMPORTANT]
> The embedded workbook and the packaged add-in are different artifacts and have
> failed differently in this project's history, and a packaged run has found
> defects an embedded run did not. An unqualified "PASS" that does not name its
> host is not evidence.

---

## 7. Manual validation

Record what you exercised and what happened. "Worked" is not a result.

| Scenario | Result |
| --- | --- |
| `DP_Show` / `DP_Close` in a clean session | <!-- --> |
| Single-cell write-back | <!-- --> |
| Table-column fill | <!-- --> |
| Formula preservation | <!-- --> |
| Protected sheet / array-formula refusal | <!-- --> |
| Each enabled entry path (keyboard, context menu, grid icon, Ribbon) | <!-- --> |
| Packaged Ribbon visuals/resources (not merely callback execution) | <!-- --> |
| Second-provider refusal | <!-- --> |
| `DP_Stop` leaves no registration behind | <!-- --> |

Checks the regression pack cannot make, because they need a real Excel session
or a real file:

```text
<!-- e.g. live OnTime delivery; workbook rename requalification; install path -->
```

---

## 8. Artifacts

| Artifact | Filename | SHA-256 | Size |
| --- | --- | --- | --- |
| Add-in | <!-- --> | <!-- --> | <!-- --> |
| Demo workbook | <!-- --> | <!-- --> | <!-- --> |

Built from SHA: <!-- must equal the reviewed source SHA above -->

Package-input inventory:

```text
VBE/VBA source:
<!-- .bas/.cls/.frm + required .frx -->

Open XML/Ribbon source:
<!-- customUI14.xml, relationship/mapping inputs, custom images/resources; N/A only if no Ribbon ships -->
```

```text
certutil -hashfile "<filename>" SHA256
```

### Source-identity discipline

- [ ] No candidate source changed after these artifacts were built.
- [ ] If it did: artifacts were rebuilt from the new SHA, package-tested again,
      and re-hashed — and the prior package-test and hash evidence was discarded.
- [ ] Nothing was rebuilt between tagging and upload.

> [!CAUTION]
> A rebuild after hashing silently invalidates the published digest. The file
> people download then does not match the hash they are told to verify, and
> nothing in the release surfaces the mismatch.

---

## 9. Deviations

Anything that did not go to plan, was skipped, or was accepted as a known
limitation. Write it here rather than leaving it out.

```text
<!-- none, or an itemized list with rationale and owning issue -->
```

---

## 10. Post-publication verification

Performed after the release is public, by re-reading GitHub rather than trusting
what was uploaded.

| Check | Result |
| --- | --- |
| `git rev-parse vX.Y.Z` matches the recorded tag object SHA | <!-- --> |
| `git rev-parse vX.Y.Z^{commit}` matches the recorded tag target SHA | <!-- --> |
| Release points at that tag | <!-- --> |
| `VERSION` and the dated changelog section match the tag | <!-- --> |
| Asset filenames match exactly what was hashed | <!-- --> |
| Each asset downloads, and its SHA-256 matches the published digest | <!-- --> |
| Packaged artifact opens from the downloaded copy and passes its smoke test | <!-- --> |
| Source archive contains expected VBA and Open XML/Ribbon source inputs | <!-- --> |
| Installation links and examples in the release notes resolve | <!-- --> |
| Default branch is ready for the next Unreleased cycle | <!-- --> |

Re-hash the **downloaded** files, not the local build outputs:

```text
certutil -hashfile "<downloaded filename>" SHA256
```

> [!CAUTION]
> Never delete and recreate a public tag to correct a mistake. Anyone who already
> fetched it keeps the old object. Correct forward instead, and record the
> correction here.

---

## 11. Provenance limits

State these plainly in the release rather than letting a reader infer stronger
guarantees than exist.

- [ ] No CI runs the regression pack. Every figure in this record was produced by
      a human on a named host
      ([#15](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/15) remains
      open).
- [ ] Artifact hashes establish file identity only. They are not cryptographic
      source-to-binary provenance: nothing here proves the published binary was
      built from the recorded source
      ([#16](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/16) remains
      open).
- [ ] The `Public` surface is larger than the supported API, and formal
      classification is outstanding
      ([#25](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/25) remains
      open).

---

## 12. Sign-off

- [ ] Every field above is filled or explicitly marked N/A with a reason.
- [ ] Every figure was pasted from a run, not transcribed.
- [ ] Both hosts are recorded.
- [ ] Artifact hashes were computed from the exact published files.
- [ ] Deviations are recorded, including any this certification chose to accept.
- [ ] The required order was followed, or every departure is recorded.
- [ ] Post-publication verification was performed against the live release.

**Certified:** <!-- name, date -->

---

## Related

- [RELEASING.md](https://github.com/danielep71/VBA-DATETIMEPICKER/blob/main/RELEASING.md) — the procedure this record documents
- [CONTRIBUTING.md](https://github.com/danielep71/VBA-DATETIMEPICKER/blob/main/CONTRIBUTING.md) — engineering contracts
- [INSTALLATION.md](https://github.com/danielep71/VBA-DATETIMEPICKER/blob/main/INSTALLATION.md) — install and validation paths
