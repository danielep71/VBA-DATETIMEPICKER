"""Fixtures for build_ribbon_package.py: Ribbon injection into a minimal Office package."""
import io
import json
import shutil
import tempfile
import unittest
import zipfile
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

import build_ribbon_package as builder

RIBBON_ROOT = Path(__file__).resolve().parent.parent / "src" / "ribbon"
CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.addin.macroEnabled.main+xml"/>'
    "</Types>"
)
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="xl/workbook.xml"/></Relationships>'
)
WORKBOOK = b"<workbook/>"


def make_package(folder: Path, extra: dict[str, bytes] | None = None,
                 content_types: str = CONTENT_TYPES, root_rels: str = ROOT_RELS) -> Path:
    package = folder / "fixture.xlam"
    with zipfile.ZipFile(package, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", WORKBOOK)
        archive.writestr("xl/vbaProject.bin", b"\xd0\xcf\x11\xe0 not a real project")
        for name, payload in (extra or {}).items():
            archive.writestr(name, payload)
    return package


def copy_ribbon(folder: Path) -> Path:
    target = folder / "ribbon"
    shutil.copytree(RIBBON_ROOT, target)
    return target


class InjectionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.folder)

    def test_parts_are_byte_identical_to_the_source(self) -> None:
        package = make_package(self.folder)
        report = builder.inject(package, RIBBON_ROOT)
        with zipfile.ZipFile(package) as archive:
            for source, part in (
                ("customUI.xml", "customUI/customUI.xml"),
                ("_rels/customUI.xml.rels", "customUI/_rels/customUI.xml.rels"),
                ("images/dp_show_picker.png", "customUI/images/dp_show_picker.png"),
                ("images/dp_reset.png", "customUI/images/dp_reset.png"),
                ("images/dp_demo.png", "customUI/images/dp_demo.png"),
            ):
                self.assertEqual(archive.read(part), (RIBBON_ROOT / source).read_bytes(), part)
        self.assertEqual(len(report["parts"]), 5)
        self.assertEqual(builder.verify(package, RIBBON_ROOT), [])

    def test_root_relationship_is_added_and_other_entries_are_unchanged(self) -> None:
        package = make_package(self.folder)
        builder.inject(package, RIBBON_ROOT)
        with zipfile.ZipFile(package) as archive:
            rels = builder.parse_xml(archive.read("_rels/.rels"), "rels")
            types = [element.get("Type") for element in rels]
            self.assertEqual(types.count(builder.EXTENSIBILITY_TYPE), 1)
            ribbon = [element for element in rels if element.get("Type") == builder.EXTENSIBILITY_TYPE][0]
            self.assertEqual(ribbon.get("Target"), "customUI/customUI.xml")
            self.assertEqual(ribbon.get("Id"), builder.RELATIONSHIP_ID)
            self.assertIn('Target="xl/workbook.xml"', archive.read("_rels/.rels").decode("utf-8"))
            self.assertEqual(archive.read("xl/workbook.xml"), WORKBOOK)
            self.assertEqual(archive.namelist()[:4],
                             ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/vbaProject.bin"])

    def test_png_content_type_is_added_once(self) -> None:
        package = make_package(self.folder)
        report = builder.inject(package, RIBBON_ROOT)
        self.assertEqual(report["content_type_defaults_added"], ["png"])
        with zipfile.ZipFile(package) as archive:
            text = archive.read("[Content_Types].xml").decode("utf-8")
        self.assertEqual(text.count('Extension="png"'), 1)
        self.assertIn('<Default Extension="png" ContentType="image/png"/>', text)

    def test_existing_png_default_is_kept(self) -> None:
        types = CONTENT_TYPES.replace("</Types>", '<Default Extension="PNG" ContentType="image/png"/></Types>')
        package = make_package(self.folder, content_types=types)
        report = builder.inject(package, RIBBON_ROOT)
        self.assertEqual(report["content_type_defaults_added"], [])

    def test_existing_png_default_with_another_value_is_kept(self) -> None:
        # The shipped v1.2.2 add-in declares PNG as image/.PNG, and Excel loads it.
        types = CONTENT_TYPES.replace("</Types>", '<Default Extension="png" ContentType="image/.PNG"/></Types>')
        package = make_package(self.folder, content_types=types)
        builder.inject(package, RIBBON_ROOT)
        with zipfile.ZipFile(package) as archive:
            text = archive.read("[Content_Types].xml").decode("utf-8")
        self.assertIn('ContentType="image/.PNG"', text)
        self.assertEqual(text.count('Extension="png"'), 1)
        self.assertEqual(builder.verify(package, RIBBON_ROOT), [])

    def test_relationship_id_does_not_collide(self) -> None:
        rels = ROOT_RELS.replace("</Relationships>", f'<Relationship Id="{builder.RELATIONSHIP_ID}" '
                                 'Type="urn:other" Target="docProps/app.xml"/></Relationships>')
        package = make_package(self.folder, root_rels=rels)
        report = builder.inject(package, RIBBON_ROOT)
        self.assertEqual(report["root_relationship"]["id"], builder.RELATIONSHIP_ID + "2")

    def test_package_with_ribbon_parts_is_refused(self) -> None:
        package = make_package(self.folder, {"customUI/customUI.xml": b"<customUI/>"})
        before = package.read_bytes()
        with self.assertRaisesRegex(builder.PackageError, "already has Ribbon parts"):
            builder.inject(package, RIBBON_ROOT)
        self.assertEqual(package.read_bytes(), before)

    def test_package_with_extensibility_relationship_is_refused(self) -> None:
        rels = ROOT_RELS.replace("</Relationships>", '<Relationship Id="r9" '
                                 f'Type="{builder.EXTENSIBILITY_TYPES[1]}" Target="customUI/customUI14.xml"/>'
                                 "</Relationships>")
        package = make_package(self.folder, root_rels=rels)
        with self.assertRaisesRegex(builder.PackageError, "extensibility"):
            builder.inject(package, RIBBON_ROOT)

    def test_missing_image_is_refused(self) -> None:
        ribbon = copy_ribbon(self.folder)
        (ribbon / "images" / "dp_reset.png").unlink()
        package = make_package(self.folder)
        before = package.read_bytes()
        with self.assertRaisesRegex(builder.PackageError, "missing file: images/dp_reset.png"):
            builder.inject(package, ribbon)
        self.assertEqual(package.read_bytes(), before)

    def test_image_without_relationship_is_refused(self) -> None:
        ribbon = copy_ribbon(self.folder)
        part = ribbon / "customUI.xml"
        part.write_bytes(part.read_bytes().replace(b'image="dp_demo"', b'image="ghost"'))
        with self.assertRaisesRegex(builder.PackageError, "no relationship: ghost"):
            builder.inject(make_package(self.folder), ribbon)

    def test_target_outside_the_ribbon_root_is_refused(self) -> None:
        ribbon = copy_ribbon(self.folder)
        rels = ribbon / "_rels" / "customUI.xml.rels"
        rels.write_bytes(rels.read_bytes().replace(b'Target="images/dp_demo.png"', b'Target="../dp_demo.png"'))
        with self.assertRaisesRegex(builder.PackageError, "outside the Ribbon root"):
            builder.inject(make_package(self.folder), ribbon)

    def test_document_type_declaration_is_refused(self) -> None:
        ribbon = copy_ribbon(self.folder)
        part = ribbon / "customUI.xml"
        part.write_bytes(b'<!DOCTYPE x [<!ENTITY a "b">]>' + part.read_bytes())
        with self.assertRaisesRegex(builder.PackageError, "document type declaration"):
            builder.inject(make_package(self.folder), ribbon)

    def test_verify_reports_a_changed_part(self) -> None:
        package = make_package(self.folder)
        builder.inject(package, RIBBON_ROOT)
        tampered = self.folder / "tampered.xlam"
        with zipfile.ZipFile(package) as source, zipfile.ZipFile(tampered, "w") as target:
            for entry in source.infolist():
                payload = source.read(entry)
                if entry.filename == "customUI/images/dp_demo.png":
                    payload = payload + b"\0"
                target.writestr(entry, payload)
        self.assertEqual(builder.verify(tampered, RIBBON_ROOT), ["customUI/images/dp_demo.png differs from its source"])

    def test_cli_writes_the_report_and_fails_closed(self) -> None:
        package = make_package(self.folder)
        report = self.folder / "report.json"
        with redirect_stdout(io.StringIO()) as out:
            self.assertEqual(builder.main(["--package", str(package), "--report", str(report)]), 0)
        self.assertIn("PASS", out.getvalue())
        self.assertEqual(json.loads(report.read_text(encoding="utf-8"))["root_relationship"]["target"],
                         "customUI/customUI.xml")
        with redirect_stderr(io.StringIO()) as err:
            self.assertEqual(builder.main(["--package", str(package)]), 1)
        self.assertIn("FAIL", err.getvalue())


if __name__ == "__main__":
    unittest.main()
