# dist/

This folder is where the packaged build output of the add-in is written locally.
It is empty in a fresh clone.

The add-in is generated from the VBA source under [`../src`](../src) and is
intentionally **not tracked in git** — it's a binary build artifact, so it's
excluded via `.gitignore`.

To get the add-in, download the latest release asset from the
[Releases page](https://github.com/danielep71/VBA-DATETIMEPICKER/releases). The
demo workbook is published there too. At `v1.2.1` the assets are:

```text
DATETIMEPICKER v1.2.1.xlam
DATETIMEPICKER-demo-v1.2.1.xlsm
```

The `.xlam` separator has varied between releases, so check the exact name on the
Release page rather than assuming it. Each release lists a SHA-256 for its
assets — verify it before enabling macros.
