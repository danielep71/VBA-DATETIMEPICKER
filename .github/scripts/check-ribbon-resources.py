#!/usr/bin/env python3
"""Check that the tracked RibbonX source is complete enough to rebuild the package.

src/ribbon/ mirrors the customUI/ folder inside the Office package:

    customUI/customUI.xml               -> src/ribbon/customUI.xml
    customUI/_rels/customUI.xml.rels    -> src/ribbon/_rels/customUI.xml.rels
    customUI/images/<file>              -> src/ribbon/images/<file>

For every Ribbon part (customUI.xml or customUI14.xml) found under the ribbon
root, the check fails when:

- the part is not well-formed XML;
- its namespace does not match its file name (customUI.xml is the Office 2007
  format, 2006/01; customUI14.xml is the Office 2010 format, 2009/07);
- it references a custom image (image="...") with no image relationship of
  that Id in its _rels/<part>.rels file;
- an image relationship points outside the ribbon root, or at a file Git does
  not track (an untracked local file would not reach a clean checkout);
- the relationship file contains an image relationship nothing references.

Built-in Office icons (imageMso="...") need no resource and are ignored.

Usage:
    python3 .github/scripts/check-ribbon-resources.py [--root src/ribbon]
    python3 .github/scripts/check-ribbon-resources.py --self-test
"""

import argparse
import os
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET

PART_NAMESPACES = {
    "customUI.xml": "http://schemas.microsoft.com/office/2006/01/customui",
    "customUI14.xml": "http://schemas.microsoft.com/office/2009/07/customui",
}
RELS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
IMAGE_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)


def tracked_files(root):
    """Return the set of Git-tracked paths under root, relative to root."""
    result = subprocess.run(
        ["git", "-C", root, "ls-files", "-z", "--", "."],
        capture_output=True, check=True)
    return {os.path.normpath(p) for p in result.stdout.decode("utf-8").split("\0") if p}


def check(root, tracked=None):
    """Return a list of error strings for the ribbon source under root.

    tracked is the set of Git-tracked paths relative to root; when omitted it is
    read from Git, so only files a clean checkout would contain count.
    """
    errors = []
    if tracked is None:
        try:
            tracked = tracked_files(root)
        except (OSError, subprocess.CalledProcessError) as exc:
            return [f"cannot list Git-tracked files under {root}: {exc}"]
    parts = [name for name in PART_NAMESPACES if os.path.isfile(os.path.join(root, name))]
    if not parts:
        return ["no Ribbon part (customUI.xml or customUI14.xml) under " + root]

    for part in parts:
        part_path = os.path.join(root, part)
        try:
            tree = ET.parse(part_path)
        except ET.ParseError as exc:
            errors.append(f"{part}: not well-formed XML: {exc}")
            continue

        tag = tree.getroot().tag
        namespace = tag[1:].split("}")[0] if tag.startswith("{") else ""
        if namespace != PART_NAMESPACES[part]:
            errors.append(
                f"{part}: namespace {namespace or '(none)'} does not match the file "
                f"name; expected {PART_NAMESPACES[part]}"
            )

        referenced = set()
        for element in tree.iter():
            image_id = element.get("image")
            if image_id:
                referenced.add(image_id)

        rels_path = os.path.join(root, "_rels", part + ".rels")
        targets = {}
        if os.path.isfile(rels_path):
            try:
                rels = ET.parse(rels_path).getroot()
            except ET.ParseError as exc:
                errors.append(f"_rels/{part}.rels: not well-formed XML: {exc}")
                continue
            for rel in rels.iter("{%s}Relationship" % RELS_NAMESPACE):
                if rel.get("Type") == IMAGE_REL_TYPE:
                    targets[rel.get("Id")] = rel.get("Target") or ""
        elif referenced:
            errors.append(
                f"{part}: references custom images {sorted(referenced)} but "
                f"_rels/{part}.rels is missing"
            )
            continue

        for image_id in sorted(referenced):
            if image_id not in targets:
                errors.append(
                    f"{part}: image=\"{image_id}\" has no image relationship in "
                    f"_rels/{part}.rels"
                )

        for image_id, target in sorted(targets.items()):
            relative = os.path.normpath(target)
            if os.path.isabs(target) or relative == ".." or relative.startswith(".." + os.sep):
                errors.append(
                    f"_rels/{part}.rels: Id=\"{image_id}\" targets {target}, "
                    f"which is outside {root}"
                )
            elif relative not in tracked:
                errors.append(
                    f"_rels/{part}.rels: Id=\"{image_id}\" targets {target}, "
                    f"which is not tracked under {root}"
                )
            if image_id not in referenced:
                errors.append(
                    f"_rels/{part}.rels: Id=\"{image_id}\" is not referenced by {part}"
                )

    return errors


def _write(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(text)


def self_test():
    """Exercise the passing case and each failure the check exists to catch."""
    part = (
        '<customUI xmlns="%s"><ribbon><tabs><tab idMso="TabHome">'
        '<group id="g" label="G"><button id="b" image="icon" imageMso="Copy"/>'
        "</group></tab></tabs></ribbon></customUI>"
    )
    rels = (
        '<Relationships xmlns="%s"><Relationship Id="icon" Type="%s" '
        'Target="images/icon.png"/></Relationships>' % (RELS_NAMESPACE, IMAGE_REL_TYPE)
    )
    cases = []

    def case(name, build, expect_fragment):
        cases.append((name, build, expect_fragment))

    def valid(root):
        _write(os.path.join(root, "customUI.xml"), part % PART_NAMESPACES["customUI.xml"])
        _write(os.path.join(root, "_rels", "customUI.xml.rels"), rels)
        _write(os.path.join(root, "images", "icon.png"), "png")

    case("complete source passes", valid, None)
    case("missing image file fails",
         lambda r: (valid(r), os.remove(os.path.join(r, "images", "icon.png"))),
         "not tracked")
    case("untracked image file fails",
         lambda r: (valid(r), _write(os.path.join(r, ".untracked"), "images/icon.png")),
         "not tracked")
    case("target outside the ribbon root fails",
         lambda r: (valid(r), _write(os.path.join(r, "_rels", "customUI.xml.rels"),
                    rels.replace('Target="images/icon.png"', 'Target="../../README.md"'))),
         "outside")
    case("missing relationship file fails",
         lambda r: (valid(r), os.remove(os.path.join(r, "_rels", "customUI.xml.rels"))),
         "is missing")
    case("unmapped image id fails",
         lambda r: (valid(r), _write(os.path.join(r, "customUI.xml"),
                    (part % PART_NAMESPACES["customUI.xml"]).replace('image="icon"', 'image="other"'))),
         "has no image relationship")
    case("namespace mismatch fails",
         lambda r: (valid(r), _write(os.path.join(r, "customUI.xml"),
                    part % PART_NAMESPACES["customUI14.xml"])),
         "does not match the file name")
    case("no ribbon part fails", lambda r: None, "no Ribbon part")

    failures = 0
    for name, build, expect in cases:
        with tempfile.TemporaryDirectory() as root:
            build(root)
            # Every file written counts as tracked, except the image a case marks
            # as untracked by writing its path into ".untracked"
            tracked = set()
            for folder, _, files in os.walk(root):
                for filename in files:
                    tracked.add(os.path.normpath(
                        os.path.relpath(os.path.join(folder, filename), root)))
            marker = os.path.join(root, ".untracked")
            if os.path.isfile(marker):
                with open(marker, encoding="utf-8") as handle:
                    tracked.discard(os.path.normpath(handle.read().strip()))
            errors = check(root, tracked)
        if expect is None:
            ok = not errors
        else:
            ok = any(expect in error for error in errors)
        print(("PASS " if ok else "FAIL ") + name + ("" if ok else f" -> {errors}"))
        failures += 0 if ok else 1
    return failures


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", default=os.path.join("src", "ribbon"))
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()

    if args.self_test:
        return 1 if self_test() else 0

    errors = check(args.root)
    for error in errors:
        print("::error::" + error if os.environ.get("GITHUB_ACTIONS") else "ERROR " + error)
    if not errors:
        print(f"Ribbon resources complete under {args.root}")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
