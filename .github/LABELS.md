# Repository labels and automation

The 20 core label names, colors and descriptions follow
[EXCEL-VBA-PROJECT-TEMPLATE](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE)
at `b903fe44ef6a032c4689870b83745afa1c22490d`.
The `ui-component` overlay preserves `code`, `wiki` and `winapi`.

Edit `labels.json` to change the catalogue. `labels-policy.json` selects the
label overlay only; it does not claim full template profile certification.
The sync and drift scripts originate from that template revision. The local
adaptation imports the reconciler's structured validation and planning API for
drift verification; Markdown summaries are display output only. Importing the
reconciler does not run its CLI. Delimiter fixtures cover pipe/backtick names,
clean and drift outcomes, invalid input and canonical/local disagreement.

- Pull requests validate policy and run offline reconciliation/drift fixtures.
- A push to main reconciles live labels and verifies the result. Manual sync
  is also available. Metadata updates preserve existing issue assignments.
- Daily drift detection is read-only and uploads its comparison as an artifact.
- `prune: true` removes labels absent from the selected catalogue. Add any
  new project labels to the overlay before merging; review the plan first.
- Weekly Dependabot PRs maintain pinned GitHub Actions versions.
- Traffic export retains its analytics environment and history branch, with
  write permissions scoped to its job. It does not run repository test code.

Local checks:

```sh
node .github/scripts/labels-sync.mjs --policy .github/labels-policy.json --self-test
node .github/scripts/labels-drift.mjs --policy .github/labels-policy.json --self-test
node .github/scripts/labels-sync.mjs --policy .github/labels-policy.json --mode plan --live labels-live.json
```

`labels-live.json` is a downloaded GitHub labels API array, including all pages.
DateTimePicker source and release verification are documented in
[tools/README.md](../tools/README.md). Full-template provisioning, wiki and
portfolio workflows are not adopted.
