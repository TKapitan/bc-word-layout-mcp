"""Tests for tools/scrub_corpus_metadata.py.

Run from the repository root:  python -m unittest discover -s tools/tests -v

Every test works on synthetic packages built in a temporary directory, or on COPIES of
tests/corpus/*.docx: nothing under tests/corpus/ is ever opened for writing.
"""

from __future__ import annotations

import dataclasses
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from dataclasses import dataclass
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent.parent
SCRIPT = TOOLS_DIR / "scrub_corpus_metadata.py"
CORPUS_DIR = TOOLS_DIR.parent / "tests" / "corpus"
sys.path.insert(0, str(TOOLS_DIR))

import scrub_corpus_metadata as scrub  # noqa: E402

OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
BIBLIOGRAPHY_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
DATASET_NS = "urn:microsoft-dynamics-nav/reports/Test_Report/50000/"
PROPS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
ENTRY_TIME = (2001, 2, 3, 4, 5, 6)  # an old, fixed stamp, so a rewrite that loses it shows

BIBLIOGRAPHY = (
    '<?xml version="1.0" standalone="no"?>'
    '<b:Sources SelectedStyle="\\APASixthEditionOfficeOnline.xsl" StyleName="APA" Version="6"'
    f' xmlns:b="{BIBLIOGRAPHY_NS}" xmlns="{BIBLIOGRAPHY_NS}"></b:Sources>').encode()

# A whole Word document pasted into a dataset sample field (rule 6), as it appears escaped.
FLAT_OPC_SAMPLE = (
    '&lt;?xml version="1.0" standalone="yes"?&gt;'
    '&lt;pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"&gt;'
    '&lt;/pkg:package&gt;')


def dataset(notes: str = "Notes", encoding: str = "utf-16") -> bytes:
    """A BC dataset part whose Notes column holds the given sample value."""
    text = (f'<?xml version="1.0" encoding="{encoding}"?>\r\n'
            f'<NavWordReportXmlPart xmlns="{DATASET_NS}">\r\n'
            '  <Header>\r\n'
            '    <No>No</No>\r\n'
            f'    <Notes>{notes}</Notes>\r\n'
            '  </Header>\r\n'
            '</NavWordReportXmlPart>')
    if encoding == "utf-16":
        return b"\xff\xfe" + text.encode("utf-16-le")
    return b"\xef\xbb\xbf" + text.encode("utf-8")


def core_props(creator: str = "", modified_by: str = "", revision: str = "1") -> bytes:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<cp:coreProperties'
            ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
            ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
            ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title></dc:title><dc:creator>{creator}</dc:creator>'
            f'<cp:lastModifiedBy>{modified_by}</cp:lastModifiedBy>'
            f'<cp:revision>{revision}</cp:revision></cp:coreProperties>').encode()


def app_props(company: str = "", manager: str | None = None,
              hyperlink_base: str | None = None) -> bytes:
    optional = ""
    if manager is not None:
        optional += f"<Manager>{manager}</Manager>"
    if hyperlink_base is not None:
        optional += f"<HyperlinkBase>{hyperlink_base}</HyperlinkBase>"
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Properties'
            ' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
            ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            f'<Template>Normal.dotm</Template><Company>{company}</Company>{optional}'
            '<AppVersion>16.0000</AppVersion></Properties>').encode()


@dataclass(frozen=True)
class RootPart:
    """A part the package root references from _rels/.rels."""
    name: str
    rel_type: str
    content_type: str
    payload: bytes
    target: str = ""  # the relationship target, when it is not just the part name


CORE = RootPart("docProps/core.xml", PACKAGE_REL + "/metadata/core-properties",
                "application/vnd.openxmlformats-package.core-properties+xml", core_props())
APP = RootPart("docProps/app.xml", OFFICE_REL + "/extended-properties",
               "application/vnd.openxmlformats-officedocument.extended-properties+xml",
               app_props())
CUSTOM = RootPart(
    "docProps/custom.xml", OFFICE_REL + "/custom-properties",
    "application/vnd.openxmlformats-officedocument.custom-properties+xml",
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"'
    b' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    b'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MSIP_Label_Owner">'
    b'<vt:lpwstr>someone@example.invalid</vt:lpwstr></property></Properties>')
LABEL_INFO = RootPart(
    "docMetadata/LabelInfo.xml",
    "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels",
    "application/vnd.ms-office.classificationlabels+xml",
    b'<?xml version="1.0" encoding="utf-8" standalone="yes"?>'
    b'<clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata">'
    b'<clbl:label id="{00000000-0000-0000-0000-000000000001}" enabled="1" method="Standard"'
    b' siteId="{00000000-0000-0000-0000-000000000002}" removed="0" /></clbl:labelList>')


@dataclass(frozen=True)
class Item:
    """A custom XML part, referenced from word/document.xml like Word and BC reference them."""
    index: int
    payload: bytes
    folder: str = "customXml"  # real exports spell it customXml/ or customXML/
    absolute: bool = False     # document rel target "/customXml/itemN.xml", not "../customXml/..."
    props: bool = True         # itemProps part + item rels + Override (BC-served parts have none)
    target: str = ""           # explicit document rel target, overriding the two forms above
    override: str = ""         # explicit Override PartName of the itemProps part


def rels_xml(relationships: list[tuple[str, str, str]]) -> bytes:
    body = "".join(f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"/>'
                   for rel_id, rel_type, target in relationships)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<Relationships xmlns="{PACKAGE_REL}">{body}</Relationships>').encode()


def build_docx(path: Path, *, root_parts: tuple[RootPart, ...] = (CORE, APP),
               items: tuple[Item, ...] = (), settings_rels: bytes | None = None,
               extra_overrides: tuple[tuple[str, str], ...] = ()) -> None:
    """Write a minimal .docx carrying just the plumbing the scrubber reads."""
    overrides = [("/word/document.xml", "application/vnd.openxmlformats-officedocument"
                                        ".wordprocessingml.document.main+xml")]
    root_rels = [("rId1", OFFICE_REL + "/officeDocument", "word/document.xml")]
    doc_rels: list[tuple[str, str, str]] = []
    parts: list[tuple[str, bytes]] = []
    for n, part in enumerate(root_parts, start=2):
        root_rels.append((f"rId{n}", part.rel_type, part.target or part.name))
        overrides.append(("/" + part.name, part.content_type))
        parts.append((part.name, part.payload))
    for item in items:
        folder, i = item.folder, item.index
        target = item.target or (f"/{folder}/item{i}.xml" if item.absolute
                                 else f"../{folder}/item{i}.xml")
        doc_rels.append((f"rIdItem{i}", OFFICE_REL + "/customXml", target))
        parts.append((f"{folder}/item{i}.xml", item.payload))
        if item.props:
            schema = BIBLIOGRAPHY_NS if item.payload == BIBLIOGRAPHY else DATASET_NS
            parts.append((f"{folder}/itemProps{i}.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\r\n'
                f'<ds:datastoreItem ds:itemID="{{00000000-0000-0000-0000-{i:012d}}}"'
                ' xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
                f'<ds:schemaRefs><ds:schemaRef ds:uri="{schema}"/></ds:schemaRefs>'
                '</ds:datastoreItem>').encode()))
            parts.append((f"{folder}/_rels/item{i}.xml.rels", rels_xml(
                [("rId1", OFFICE_REL + "/customXmlProps", f"itemProps{i}.xml")])))
            overrides.append((item.override or f"/{folder}/itemProps{i}.xml", PROPS_CONTENT_TYPE))
    if settings_rels is not None:
        doc_rels.append(("rIdSettings", OFFICE_REL + "/settings", "settings.xml"))
        overrides.append(("/word/settings.xml", "application/vnd.openxmlformats-officedocument"
                                                ".wordprocessingml.settings+xml"))
        parts.append(("word/settings.xml",
                      b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml'
                      b'/2006/main"><w:attachedTemplate r:id="rId1" xmlns:r="' + OFFICE_REL.encode()
                      + b'"/></w:settings>'))
        parts.append(("word/_rels/settings.xml.rels", settings_rels))
    overrides += extra_overrides
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package'
        '.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
        + "".join(f'<Override PartName="{name}" ContentType="{ctype}"/>'
                  for name, ctype in overrides)
        + '</Types>').encode()
    entries = [("[Content_Types].xml", content_types), ("_rels/.rels", rels_xml(root_rels)),
               ("word/document.xml",
                b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006'
                b'/main"><w:body><w:p/></w:body></w:document>'),
               ("word/_rels/document.xml.rels", rels_xml(doc_rels))] + parts
    with zipfile.ZipFile(path, "w") as zf:
        for name, payload in entries:
            info = zipfile.ZipInfo(name, date_time=ENTRY_TIME)
            # One stored entry, so per-entry compression is visibly preserved too.
            info.compress_type = (zipfile.ZIP_STORED if name == "word/document.xml"
                                  else zipfile.ZIP_DEFLATED)
            zf.writestr(info, payload)


def read_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


class ScrubCase(unittest.TestCase):
    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)

    def build(self, name: str = "layout.docx", **options) -> Path:
        path = self.tmp / name
        build_docx(path, **options)
        return path

    def scrub(self, path: Path) -> list[str]:
        """Scrub in write mode, and prove a second run has nothing left to do."""
        actions = scrub.scrub_file(path, check_only=False)
        self.assertEqual(scrub.scrub_file(path, check_only=True), [], "second run is not a no-op")
        return actions

    def run_cli(self, *args: object) -> subprocess.CompletedProcess:
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        return subprocess.run([sys.executable, str(SCRIPT), *map(str, args)],
                              capture_output=True, encoding="utf-8", env=env, timeout=120,
                              check=False)

    def assert_no_trace_of(self, path: Path, *part_names: str) -> None:
        """Neither the parts nor any relationship or Override naming them remain (any casing)."""
        parts = read_parts(path)
        names = {name.lower() for name in parts}
        plumbing = b"".join(data for name, data in parts.items()
                            if name.endswith(".rels") or name == "[Content_Types].xml").lower()
        for part in part_names:
            self.assertNotIn(part.lower(), names)
            self.assertNotIn(part.rpartition("/")[2].lower().encode(), plumbing, part)


class BibliographyTests(ScrubCase):
    def test_customXml_folder_with_relative_targets(self):
        path = self.build(items=(Item(1, BIBLIOGRAPHY), Item(2, dataset())))
        before = read_parts(path)

        actions = self.scrub(path)

        self.assertEqual(actions, ["dropped customXml/_rels/item1.xml.rels",
                                   "dropped customXml/item1.xml",
                                   "dropped customXml/itemProps1.xml"])
        self.assert_no_trace_of(path, "customXml/item1.xml", "customXml/itemProps1.xml",
                                "customXml/_rels/item1.xml.rels")
        after = read_parts(path)
        for name in ("customXml/item2.xml", "customXml/itemProps2.xml",
                     "customXml/_rels/item2.xml.rels", "word/document.xml"):
            self.assertEqual(after[name], before[name], name)
        self.assertIn(b'Target="../customXml/item2.xml"', after["word/_rels/document.xml.rels"])
        self.assertIn(b'PartName="/customXml/itemProps2.xml"', after["[Content_Types].xml"])

    def test_customXML_folder_with_absolute_targets(self):
        """The StandardStatement.docx shape: customXML/ and absolute document rel targets, next
        to a BC dataset part that has no itemProps at all."""
        path = self.build(items=(Item(1, BIBLIOGRAPHY, folder="customXML", absolute=True),
                                 Item(3, dataset(), folder="customXML", absolute=True,
                                      props=False)))
        before = read_parts(path)

        actions = self.scrub(path)

        self.assertEqual(actions, ["dropped customXML/_rels/item1.xml.rels",
                                   "dropped customXML/item1.xml",
                                   "dropped customXML/itemProps1.xml"])
        self.assert_no_trace_of(path, "customXML/item1.xml", "customXML/itemProps1.xml",
                                "customXML/_rels/item1.xml.rels")
        after = read_parts(path)
        self.assertEqual(after["customXML/item3.xml"], before["customXML/item3.xml"])
        self.assertIn(b'Target="/customXML/item3.xml"', after["word/_rels/document.xml.rels"])

    def test_references_match_the_part_whatever_their_casing(self):
        path = self.build(items=(Item(1, BIBLIOGRAPHY, folder="customXML",
                                      target="/CustomXml/Item1.xml",
                                      override="/customxml/ITEMPROPS1.XML"),))

        self.scrub(path)

        self.assert_no_trace_of(path, "customXML/item1.xml", "customXML/itemProps1.xml",
                                "customXML/_rels/item1.xml.rels")

    def test_every_bibliography_part_goes_in_one_run(self):
        path = self.build(items=(Item(1, BIBLIOGRAPHY),
                                 Item(2, BIBLIOGRAPHY, folder="customXML", absolute=True),
                                 Item(3, dataset())))

        actions = scrub.scrub_file(path, check_only=False)

        self.assertEqual(len([a for a in actions if a.startswith("dropped ")]), 6, actions)
        self.assert_no_trace_of(path, "customXml/item1.xml", "customXml/itemProps1.xml",
                                "customXML/item2.xml", "customXML/itemProps2.xml")
        self.assertIn("customXml/item3.xml", read_parts(path))
        second = self.run_cli("--check", path)
        self.assertEqual(second.returncode, 0, second.stdout + second.stderr)
        self.assertIn("clean   layout.docx", second.stdout)

    def test_a_reference_the_script_cannot_resolve_is_refused(self):
        """Dropping the part would leave a dangling relationship, so nothing is written."""
        path = self.build(items=(Item(1, BIBLIOGRAPHY, target="..\\customXml\\item1.xml"),))
        before = path.read_bytes()

        with self.assertRaisesRegex(scrub.ScrubError, "no relationship points at it"):
            scrub.scrub_file(path, check_only=False)
        self.assertEqual(path.read_bytes(), before)

    def test_duplicate_overrides_are_refused(self):
        path = self.build(items=(Item(1, BIBLIOGRAPHY),),
                          extra_overrides=(("/customXml/itemProps1.xml", PROPS_CONTENT_TYPE),))

        with self.assertRaisesRegex(scrub.ScrubError, "Overrides"):
            scrub.scrub_file(path, check_only=True)

    def test_unexpected_relationship_from_a_bibliography_part_is_refused(self):
        """Its item rels points at something other than an itemProps part: dropping that too, or
        leaving it behind, would both be guesses."""
        path = self.build(items=(Item(1, BIBLIOGRAPHY),))
        parts = read_parts(path)
        parts["customXml/_rels/item1.xml.rels"] = rels_xml(
            [("rId1", OFFICE_REL + "/customXmlProps", "itemProps1.xml"),
             ("rId2", OFFICE_REL + "/styles", "../word/styles.xml")])
        with zipfile.ZipFile(path, "w") as zf:
            for name, payload in parts.items():
                zf.writestr(name, payload)

        with self.assertRaisesRegex(scrub.ScrubError, "unexpected relationship"):
            scrub.scrub_file(path, check_only=True)

    def test_entries_naming_the_same_part_are_refused(self):
        path = self.build(items=(Item(1, BIBLIOGRAPHY),
                                 Item(1, dataset(), folder="customXML", props=False)))

        with self.assertRaisesRegex(scrub.ScrubError, "name the same part"):
            scrub.scrub_file(path, check_only=True)


class DroppedPartTests(ScrubCase):
    def test_custom_properties_and_label_info_go_with_absolute_root_targets(self):
        path = self.build(root_parts=(CORE, APP,
                                      dataclasses.replace(CUSTOM, target="/docProps/custom.xml"),
                                      dataclasses.replace(LABEL_INFO,
                                                          target="/docMetadata/LabelInfo.xml")))

        actions = self.scrub(path)

        self.assertEqual(actions, ["dropped docMetadata/LabelInfo.xml",
                                   "dropped docProps/custom.xml"])
        self.assert_no_trace_of(path, "docProps/custom.xml", "docMetadata/LabelInfo.xml")
        root_rels = read_parts(path)["_rels/.rels"]
        self.assertIn(b'Target="docProps/core.xml"', root_rels)
        self.assertIn(b'Target="docProps/app.xml"', root_rels)

    def test_drop_parts_are_matched_case_insensitively(self):
        path = self.build(root_parts=(CORE, dataclasses.replace(CUSTOM, name="DocProps/Custom.xml",
                                                                target="docprops/custom.XML")))

        actions = self.scrub(path)

        self.assertEqual(actions, ["dropped DocProps/Custom.xml"])
        self.assert_no_trace_of(path, "docProps/custom.xml")


class PropertyTests(ScrubCase):
    def test_core_properties_are_blanked(self):
        dirty = core_props(creator="Jane Example", modified_by="John Example", revision="7")
        path = self.build(root_parts=(dataclasses.replace(CORE, name="docProps/Core.xml",
                                                          payload=dirty), APP))

        actions = self.scrub(path)

        self.assertEqual(actions, ["docProps/Core.xml: blanked <dc:creator>",
                                   "docProps/Core.xml: blanked <cp:lastModifiedBy>",
                                   "docProps/Core.xml: reset <cp:revision> to 1"])
        self.assertEqual(read_parts(path)["docProps/Core.xml"], core_props())

    def test_app_company_and_manager_are_blanked_and_hyperlink_base_kept(self):
        dirty = app_props(company="Example Company Ltd", manager="Jane Example",
                          hyperlink_base="https://example.invalid/layouts/")
        path = self.build(root_parts=(CORE, dataclasses.replace(APP, payload=dirty)))

        actions = self.scrub(path)

        self.assertEqual(actions, ["docProps/app.xml: blanked <Company>",
                                   "docProps/app.xml: blanked <Manager>"])
        self.assertEqual(read_parts(path)["docProps/app.xml"],
                         app_props(company="", manager="",
                                   hyperlink_base="https://example.invalid/layouts/"))

    def test_missing_core_and_app_parts_are_fine(self):
        path = self.build(root_parts=())

        self.assertEqual(scrub.scrub_file(path, check_only=True), [])
        result = self.run_cli("--check", path)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


class DatasetValueTests(ScrubCase):
    def test_flat_opc_sample_value_is_replaced_in_a_customXML_folder(self):
        for encoding in ("utf-16", "utf-8"):
            with self.subTest(encoding=encoding):
                path = self.build(f"{encoding}.docx", items=(
                    Item(3, dataset(FLAT_OPC_SAMPLE, encoding), folder="customXML",
                         absolute=True, props=False),))

                actions = self.scrub(path)

                self.assertEqual(actions, ["customXML/item3.xml: replaced flat-OPC package "
                                           "sample value in <Notes>"])
                self.assertEqual(read_parts(path)["customXML/item3.xml"],
                                 dataset("Notes", encoding))


class AttachedTemplateTests(ScrubCase):
    def test_target_is_sanitized_whatever_the_attribute_order(self):
        orders = {
            "Word order": ('Id="rId1" Type="{type}" Target="{target}" TargetMode="External"'),
            "TargetMode first": ('TargetMode="External" Target="{target}" Id="rId1" Type="{type}"'),
        }
        for label, attributes in orders.items():
            with self.subTest(label):
                def settings_rels(target: str, attributes: str = attributes) -> bytes:
                    rel = attributes.format(type=OFFICE_REL + "/attachedTemplate", target=target)
                    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                            f'<Relationships xmlns="{PACKAGE_REL}"><Relationship {rel} />'
                            '</Relationships>').encode()
                personal = "file:///C:/Users/someone/Desktop/Sales%20Template.dotx"
                path = self.build(f"{label}.docx", settings_rels=settings_rels(personal))

                actions = self.scrub(path)

                sanitized = "file:///C:/Templates/Sales%20Template.dotx"
                self.assertEqual(actions, ["word/_rels/settings.xml.rels: attachedTemplate "
                                           f"target -> {sanitized}"])
                self.assertEqual(read_parts(path)["word/_rels/settings.xml.rels"],
                                 settings_rels(sanitized))


class WriteModeTests(ScrubCase):
    def test_kept_entries_keep_order_timestamps_compression_and_bytes(self):
        path = self.build(root_parts=(dataclasses.replace(CORE, payload=core_props("Jane")), APP),
                          items=(Item(1, BIBLIOGRAPHY), Item(2, dataset())))
        with zipfile.ZipFile(path) as zf:
            before = [(i.filename, i.date_time, i.compress_type, zf.read(i))
                      for i in zf.infolist()]

        actions = self.scrub(path)

        dropped = {a[len("dropped "):] for a in actions if a.startswith("dropped ")}
        edited = {"docProps/core.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"}
        with zipfile.ZipFile(path) as zf:
            after = [(i.filename, i.date_time, i.compress_type, zf.read(i))
                     for i in zf.infolist()]
        kept = [entry for entry in before if entry[0] not in dropped]
        self.assertEqual([entry[0] for entry in after], [entry[0] for entry in kept])
        for (name, stamp, compression, payload), (_, new_stamp, new_compression, new_payload) \
                in zip(kept, after):
            self.assertEqual((new_stamp, new_compression), (stamp, compression), name)
            if name not in edited:
                self.assertEqual(new_payload, payload, name)


class CommandLineTests(ScrubCase):
    def test_clean_file_exits_0(self):
        result = self.run_cli("--check", self.build())

        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("clean   layout.docx", result.stdout)

    def test_check_exits_1_when_something_would_change_and_writes_nothing(self):
        path = self.build(items=(Item(1, BIBLIOGRAPHY, folder="customXML", absolute=True),))
        before = path.read_bytes()

        result = self.run_cli("--check", path)

        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn("would scrub layout.docx:", result.stdout)
        self.assertIn("dropped customXML/item1.xml", result.stdout)
        self.assertEqual(path.read_bytes(), before)

    def test_write_mode_exits_0_and_leaves_the_file_clean(self):
        path = self.build(items=(Item(1, BIBLIOGRAPHY),))

        result = self.run_cli(path)

        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("scrubbed layout.docx:", result.stdout)
        self.assertEqual(self.run_cli("--check", path).returncode, 0)

    def test_unknown_option_exits_2_with_a_usage_line(self):
        path = self.build(items=(Item(1, BIBLIOGRAPHY),))
        before = path.read_bytes()

        result = self.run_cli("--check", "--dry-run", path)

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("usage:", result.stderr)
        self.assertIn("--dry-run", result.stderr)
        self.assertEqual(path.read_bytes(), before)

    def test_corrupt_file_exits_2(self):
        path = self.tmp / "broken.docx"
        path.write_bytes(b"this is not a zip archive")

        result = self.run_cli("--check", path)

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("broken.docx", result.stderr)

    def test_missing_file_exits_2(self):
        result = self.run_cli("--check", self.tmp / "missing.docx")

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("not a file", result.stderr)

    def test_inconsistent_package_exits_2_and_is_not_written(self):
        path = self.build(items=(Item(1, BIBLIOGRAPHY),),
                          extra_overrides=(("/customXml/itemProps1.xml", PROPS_CONTENT_TYPE),))
        before = path.read_bytes()

        result = self.run_cli(path)

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("ERROR", result.stderr)
        self.assertEqual(path.read_bytes(), before)

    def test_a_bad_file_does_not_stop_the_others(self):
        broken = self.tmp / "broken.docx"
        broken.write_bytes(b"this is not a zip archive")
        dirty = self.build(items=(Item(1, BIBLIOGRAPHY),))

        result = self.run_cli(broken, dirty)

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("scrubbed layout.docx:", result.stdout)
        self.assertEqual(scrub.scrub_file(dirty, check_only=True), [])


class ResolveTargetTests(unittest.TestCase):
    def test_targets_resolve_against_the_source_folder_or_the_package_root(self):
        cases = [
            ("word/_rels/document.xml.rels", "../customXml/item1.xml", "customXml/item1.xml"),
            ("word/_rels/document.xml.rels", "/customXML/item3.xml", "customXML/item3.xml"),
            ("word/_rels/document.xml.rels", "styles.xml", "word/styles.xml"),
            ("word/glossary/_rels/document.xml.rels", "../../customXml/item2.xml",
             "customXml/item2.xml"),
            ("customXML/_rels/item1.xml.rels", "itemProps1.xml", "customXML/itemProps1.xml"),
            ("_rels/.rels", "docProps/custom.xml", "docProps/custom.xml"),
            ("_rels/.rels", "/docProps/custom.xml", "docProps/custom.xml"),
        ]
        for rels_name, target, expected in cases:
            with self.subTest(rels=rels_name, target=target):
                self.assertEqual(scrub.resolve_target(rels_name, target), expected)


class CorpusTests(ScrubCase):
    def test_every_corpus_file_is_already_clean(self):
        """The stricter matching finds nothing new to scrub in the real corpus (run on copies)."""
        originals = sorted(CORPUS_DIR.glob("*.docx"))
        self.assertTrue(originals, f"no corpus files under {CORPUS_DIR}")
        copies = []
        for original in originals:
            copy = self.tmp / original.name
            shutil.copyfile(original, copy)
            copies.append(copy)
        before = {copy: copy.read_bytes() for copy in copies}

        result = self.run_cli("--check", *copies)

        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        for copy in copies:
            self.assertIn(f"clean   {copy.name}", result.stdout)
            self.assertEqual(copy.read_bytes(), before[copy])


if __name__ == "__main__":
    unittest.main()
