"""Fixtures for sanitize_package.py: each kind of build-machine metadata, a clean package, and safety checks."""
import json
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

import sanitize_package as sanitizer

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
VBA = b"\xd0\xcf\x11\xe0 synthetic VBA project; never edited"
PRINTER = "Office Printer".encode("utf-16-le") + b"\x00" * 34

CLEAN_PARTS: dict[str, str | bytes] = {
    "[Content_Types].xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.addin.macroEnabled.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '</Types>'),
    "_rels/.rels": (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELS_NS}">'
        '<Relationship Id="rUI" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" '
        'Target="customUI/customUI.xml"/>'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
        'officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/'
        'core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
        'extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>'),
    "xl/workbook.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'xmlns:r="{DOC_REL}"><workbookPr codeName="ThisWorkbook"/>'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'),
    "xl/_rels/workbook.xml.rels": (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELS_NS}">'
        f'<Relationship Id="rId1" Type="{DOC_REL}/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
        'Target="vbaProject.bin"/>'
        '<Relationship Id="rLink" Type="' + DOC_REL + '/hyperlink" Target="https://example.invalid/" '
        'TargetMode="External"/>'
        '</Relationships>'),
    "xl/worksheets/sheet1.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'xmlns:r="{DOC_REL}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Keep</t></is></c></row>'
        '</sheetData><pageSetup orientation="portrait"/></worksheet>'),
    "xl/vbaProject.bin": VBA,
    "customUI/customUI.xml": '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>',
    "docProps/core.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>'
        '</cp:coreProperties>'),
    "docProps/app.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        '<Application>Microsoft Excel</Application><Company></Company></Properties>'),
}


def dirty_parts() -> dict[str, str | bytes]:
    parts = dict(CLEAN_PARTS)
    parts["[Content_Types].xml"] = str(parts["[Content_Types].xml"]).replace(
        "</Types>",
        '<Override PartName="/docMetadata/LabelInfo.xml" '
        'ContentType="application/vnd.ms-office.classificationlabels+xml"/>'
        '<Override PartName="/customXml/itemProps1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'
        '</Types>')
    parts["_rels/.rels"] = str(parts["_rels/.rels"]).replace(
        "</Relationships>",
        '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels" '
        'Target="docMetadata/LabelInfo.xml"/></Relationships>')
    parts["xl/workbook.xml"] = str(parts["xl/workbook.xml"]).replace(
        "<sheets>",
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice Requires="x15"><x15ac:absPath url="C:\\Users\\someone\\Private\\" '
        'xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice>'
        '</mc:AlternateContent><sheets>')
    parts["xl/_rels/workbook.xml.rels"] = str(parts["xl/_rels/workbook.xml.rels"]).replace(
        "</Relationships>",
        f'<Relationship Id="rId8" Type="{DOC_REL}/customXml" Target="../customXml/item1.xml"/></Relationships>')
    parts["xl/worksheets/sheet1.xml"] = str(parts["xl/worksheets/sheet1.xml"]).replace(
        '<pageSetup orientation="portrait"/>', '<pageSetup orientation="portrait" r:id="rId1"/>')
    parts["xl/worksheets/_rels/sheet1.xml.rels"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELS_NS}">'
        f'<Relationship Id="rId1" Type="{DOC_REL}/printerSettings" '
        'Target="../printerSettings/printerSettings1.bin"/></Relationships>')
    parts["xl/printerSettings/printerSettings1.bin"] = PRINTER
    parts["docMetadata/LabelInfo.xml"] = (
        '<clbl:labelList xmlns:clbl="urn:x"><clbl:label siteId="tenant"/></clbl:labelList>')
    parts["customXml/item1.xml"] = "<root>library</root>"
    parts["customXml/itemProps1.xml"] = '<ds:datastoreItem xmlns:ds="urn:x" ds:itemID="{1}"/>'
    parts["customXml/_rels/item1.xml.rels"] = (
        f'<Relationships xmlns="{RELS_NS}"><Relationship Id="rId1" '
        'Type="http://x/customXmlProps" Target="itemProps1.xml"/></Relationships>')
    parts["docProps/core.xml"] = str(parts["docProps/core.xml"]).replace(
        "<dcterms:created", "<dc:creator>Some Person</dc:creator>"
        "<cp:lastModifiedBy>Some Person</cp:lastModifiedBy><dcterms:created")
    parts["docProps/app.xml"] = str(parts["docProps/app.xml"]).replace(
        "<Company></Company>", "<Company>Some Employer</Company><Manager>Some Manager</Manager>")
    return parts


def write(path: Path, parts: dict[str, str | bytes]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content.encode("utf-8") if isinstance(content, str) else content)


def read(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


class SanitizeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.folder)
        self.dirty = self.folder / "dirty.xlam"
        write(self.dirty, dirty_parts())
        self.clean = self.folder / "clean.xlam"
        write(self.clean, CLEAN_PARTS)

    def test_every_kind_of_metadata_is_found(self) -> None:
        findings = sanitizer.check_package(self.dirty)
        for expected in ("sensitivity label", "document-library metadata", "printer settings",
                         "x15ac:absPath", "dc:creator", "cp:lastModifiedBy", "Company", "Manager"):
            self.assertTrue(any(expected in f for f in findings), f"{expected} not reported: {findings}")

    def test_clean_package_has_no_findings(self) -> None:
        self.assertEqual(sanitizer.check_package(self.clean), [])

    def test_sanitized_copy_carries_none_of_it(self) -> None:
        output = self.folder / "out.xlam"
        result = sanitizer.sanitize_package(self.dirty, output)
        self.assertEqual(sanitizer.check_package(output), [])
        parts = read(output)
        self.assertFalse([p for p in parts if p.startswith(("docMetadata/", "customXml/", "xl/printerSettings/"))])
        joined = b"".join(parts.values())
        for secret in (b"someone", b"Some Person", b"Some Employer", b"Some Manager", b"tenant"):
            self.assertNotIn(secret, joined)
        self.assertIn("xl/workbook.xml", result.edited_parts)

    def test_relationships_and_references_are_consistent(self) -> None:
        output = self.folder / "out.xlam"
        sanitizer.sanitize_package(self.dirty, output)
        parts = read(output)
        self.assertNotIn(b'r:id="rId1"', parts["xl/worksheets/sheet1.xml"].split(b"<pageSetup")[1])
        self.assertIn(b'<sheet name="Sheet1" sheetId="1" r:id="rId1"/>', parts["xl/workbook.xml"])
        self.assertNotIn(b"docMetadata", parts["_rels/.rels"] + parts["[Content_Types].xml"])
        self.assertNotIn(b"customXml", parts["xl/_rels/workbook.xml.rels"] + parts["[Content_Types].xml"])
        self.assertIn(b'TargetMode="External"', parts["xl/_rels/workbook.xml.rels"])
        self.assertIn(b"<Company></Company>", parts["docProps/app.xml"])

    def test_matching_text_in_cells_is_kept(self) -> None:
        parts = dirty_parts()
        parts["xl/worksheets/sheet1.xml"] = str(parts["xl/worksheets/sheet1.xml"]).replace(
            "<t>Keep</t>", '<t>Keep r:id="rId1" text</t>').replace(
            "</row>", '<c r="B1"><f>"r:id=""rId1"""</f></c></row>')
        write(self.dirty, parts)
        output = self.folder / "out.xlam"
        sanitizer.sanitize_package(self.dirty, output)
        sheet = read(output)["xl/worksheets/sheet1.xml"]
        self.assertIn(b'<t>Keep r:id="rId1" text</t>', sheet)
        self.assertIn(b'<f>"r:id=""rId1"""</f>', sheet)
        self.assertIn(b'<pageSetup orientation="portrait"/>', sheet)

    def test_matching_text_in_another_attribute_is_kept(self) -> None:
        parts = dirty_parts()
        parts["xl/worksheets/sheet1.xml"] = str(parts["xl/worksheets/sheet1.xml"]).replace(
            "</sheetData>", '</sheetData><hyperlinks><hyperlink ref="A1" '
            "tooltip=\"Keep r:id='rId1' text\" display='r:id=\"rId1\"'/></hyperlinks>")
        write(self.dirty, parts)
        output = self.folder / "out.xlam"
        sanitizer.sanitize_package(self.dirty, output)
        sheet = read(output)["xl/worksheets/sheet1.xml"]
        self.assertIn(b"tooltip=\"Keep r:id='rId1' text\" display='r:id=\"rId1\"'", sheet)
        self.assertIn(b'<pageSetup orientation="portrait"/>', sheet)

    def test_attribute_change_is_refused(self) -> None:
        def tamper(data: dict[str, bytes]) -> tuple[dict[str, bytes], sanitizer.Result]:
            out, result = original(data)
            out["xl/worksheets/sheet1.xml"] = out["xl/worksheets/sheet1.xml"].replace(
                b'orientation="portrait"', b'orientation="landscape"')
            return out, result

        original = sanitizer.sanitize_parts
        with mock.patch("sanitize_package.sanitize_parts", side_effect=tamper), \
                self.assertRaisesRegex(sanitizer.SanitizeError, "more than the removed"):
            sanitizer.sanitize_package(self.dirty, self.folder / "out.xlam")

    def test_text_change_is_refused(self) -> None:
        def tamper(data: dict[str, bytes]) -> tuple[dict[str, bytes], sanitizer.Result]:
            out, result = original(data)
            out["xl/worksheets/sheet1.xml"] = out["xl/worksheets/sheet1.xml"].replace(b"Keep", b"Lost")
            return out, result

        original = sanitizer.sanitize_parts
        with mock.patch("sanitize_package.sanitize_parts", side_effect=tamper), \
                self.assertRaisesRegex(sanitizer.SanitizeError, "text content changed"):
            sanitizer.sanitize_package(self.dirty, self.folder / "out.xlam")

    def test_untouched_parts_keep_their_bytes_and_order(self) -> None:
        output = self.folder / "out.xlam"
        sanitizer.sanitize_package(self.dirty, output)
        before, after = read(self.dirty), read(output)
        for part in ("xl/vbaProject.bin", "customUI/customUI.xml"):
            self.assertEqual(before[part], after[part])
        with zipfile.ZipFile(self.dirty) as a, zipfile.ZipFile(output) as b:
            self.assertEqual([n for n in a.namelist() if n in after], b.namelist())

    def test_clean_package_is_copied_unchanged(self) -> None:
        output = self.folder / "out.xlam"
        result = sanitizer.sanitize_package(self.clean, output)
        self.assertEqual((result.removed_parts, result.edited_parts), ([], []))
        self.assertEqual(read(self.clean), read(output))

    def test_in_place_replaces_the_package(self) -> None:
        with mock.patch("sys.stdout"):
            self.assertEqual(sanitizer.main([str(self.dirty), "--in-place"]), 0)
        self.assertEqual(sanitizer.check_package(self.dirty), [])
        self.assertEqual(sorted(p.name for p in self.folder.iterdir()), ["clean.xlam", "dirty.xlam"])

    def test_check_mode_exit_code_and_report(self) -> None:
        report = self.folder / "report.json"
        with mock.patch("sys.stdout"):
            self.assertEqual(sanitizer.main(["--check", str(self.dirty), str(self.clean),
                                             "--report", str(report)]), 1)
            self.assertEqual(sanitizer.main(["--check", str(self.clean)]), 0)
        data = json.loads(report.read_text(encoding="utf-8"))
        self.assertEqual([bool(p["findings"]) for p in data["packages"]], [True, False])

    def test_not_a_package_fails_without_writing(self) -> None:
        bogus = self.folder / "bogus.xlam"
        bogus.write_bytes(b"not a zip")
        with mock.patch("sys.stderr"):
            self.assertEqual(sanitizer.main([str(bogus), "--in-place"]), 2)
        self.assertEqual(bogus.read_bytes(), b"not a zip")

    def test_vba_project_change_is_refused(self) -> None:
        def tamper(data: dict[str, bytes]) -> tuple[dict[str, bytes], sanitizer.Result]:
            out, result = original(data)
            out["xl/vbaProject.bin"] = b"changed"
            return out, result

        original = sanitizer.sanitize_parts
        output = self.folder / "out.xlam"
        with mock.patch("sanitize_package.sanitize_parts", side_effect=tamper), \
                self.assertRaisesRegex(sanitizer.SanitizeError, "vbaProject.bin changed"):
            sanitizer.sanitize_package(self.dirty, output)
        self.assertFalse(output.exists())

    def test_doctype_in_an_edited_part_is_refused(self) -> None:
        parts = dirty_parts()
        body = str(parts["docProps/core.xml"]).split("?>", 1)[1]
        parts["docProps/core.xml"] = '<!DOCTYPE x [<!ENTITY a "b">]>' + body
        write(self.dirty, parts)
        with self.assertRaisesRegex(sanitizer.SanitizeError, "DOCTYPE"):
            sanitizer.sanitize_package(self.dirty, self.folder / "out.xlam")


if __name__ == "__main__":
    unittest.main()
