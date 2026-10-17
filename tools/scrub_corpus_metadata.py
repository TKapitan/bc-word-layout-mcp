"""Scrub personal metadata from the tracked corpus .docx files (release plan G1).

The corpus files are real Business Central layout captures whose *content* is cleared for
publication, but whose OOXML *metadata* names the individuals who happened to author or last
save them. This script removes exactly that residue and nothing else:

  1. docProps/core.xml      - blank dc:creator / cp:lastModifiedBy, reset cp:revision to 1
                              (the shape Microsoft's own pristine captures already have).
  2. docProps/custom.xml    - DROPPED where present. Pure MSIP sensitivity-label residue
                              (label GUIDs, tenant ids, and in one file a personal
                              microsoft.com address in the label Owner property).
  3. docMetadata/LabelInfo.xml - DROPPED where present. Same MSIP labelling, tenant GUIDs.
  4. bibliography customXml - DROPPED where present: every such part, each with its item rels
                              and itemProps parts. An empty <b:Sources> part Word inserts
                              by default; identified by its namespace, NEVER by part index,
                              because the BC dataset part shares the customXml/ folder.
                              Verified before adopting this rule: no w:dataBinding in the
                              corpus references a bibliography part's ds:itemID.
  5. attachedTemplate rels  - target REWRITTEN to file:///C:/Templates/<same filename>.
                              The originals point into named developers' Desktop folders.
                              The relationship itself is kept deliberately: it is a
                              real-world wart the validator warns about and tests exercise
                              (LayoutValidatorTests / FullValidatorTests / McpHostToolTests
                              preview-stripping); only the personal path segment goes.
  6. dataset sample values  - a leaf value in the BC dataset part that contains an escaped
                              flat-OPC package (someone pasted a whole Word document into a
                              sample field at capture time; it embeds its own LabelInfo with
                              a tenant GUID) is replaced with the element's own name, the
                              value convention every sibling element already follows. This
                              is the ONLY edit ever made inside a BC dataset part, it is
                              content-addressed (never name- or index-addressed), and the
                              part's encoding (UTF-16 LE + BOM or UTF-8) is verified to
                              round-trip byte-identically before any edit is attempted.
                              Nothing in the product reads stored dataset values (schema
                              comes from element structure; sample data is generated), and
                              Word treats the resulting storeItemChecksum mismatch as a
                              normal stale-cache refresh.
  7. docProps/app.xml       - blank Company / Manager, keeping the elements (as rule 1 does).
                              HyperlinkBase is deliberately left alone: it is a working
                              setting (what relative links resolve against), not a record of
                              who handled the file.

A dropped part takes its package plumbing with it: every relationship that targets it, from any
.rels part, and its [Content_Types].xml Override.

The plumbing is read the way OPC defines it (ECMA-376 Part 2), so no rule is skipped because of
how a package happens to spell a reference:
  - part names compare case-insensitively (real exports use both customXml/ and customXML/);
    each part keeps whatever casing the package gives it;
  - a relationship target resolves against its source part's folder, or from the package root
    when it starts with "/" (the corpus has both forms); external targets are never resolved;
  - docProps/core.xml and docProps/app.xml are optional.
A package the rules cannot be applied to cleanly is refused with an error, never half-scrubbed:
two entries naming the same part; a part to drop that no relationship references (more likely
a reference this script cannot resolve than a genuine orphan, and dropping the part would then
leave that reference dangling); duplicate Overrides; a bibliography part related to anything
but its itemProps; plumbing XML that does not parse, or that the byte-level editor cannot fully
account for.

Everything else - the BC dataset customXml part (UTF-16, byte-sensitive), the document body,
headers/footers, media, styles, every other relationship - is copied through byte-identical.
Zip entry order and per-entry timestamps are preserved.

Idempotent: one run applies every rule everywhere it matches, so a second run reports every file
clean. Run it whenever a new capture joins the corpus. Without paths it scrubs the top-level
tests/corpus/*.docx only (subfolders are private material).

Usage:  python tools/scrub_corpus_metadata.py [--check] [paths...]
        --check  report what would change; write nothing.
Exit:   0 = every file clean (or scrubbed); 1 = --check found something to scrub;
        2 = an error: a bad command line, or a package that is unreadable or inconsistent (that
            file is left untouched; the other files are still processed).
"""

from __future__ import annotations

import argparse
import html
import os
import posixpath
import re
import string
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import NoReturn

BIBLIOGRAPHY_NS = b"http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
SANITIZED_TEMPLATE_DIR = "file:///C:/Templates/"

# Parts dropped outright wherever they appear, with the package plumbing that references them.
DROP_PARTS = ("docProps/custom.xml", "docMetadata/LabelInfo.xml")
CORE_PROPS = "docProps/core.xml"
APP_PROPS = "docProps/app.xml"
CONTENT_TYPES = "[Content_Types].xml"

# Matched against zip entry names, so case-insensitively: OPC part names are case-insensitive.
CUSTOM_XML_ITEM = re.compile(r"customXml/item\d+\.xml", re.IGNORECASE)
CUSTOM_XML_PROPS = re.compile(r"customXml/itemProps\d+\.xml", re.IGNORECASE)
RELS_PART = re.compile(r"(?:.+/)?_rels/[^/]*\.rels", re.IGNORECASE)

_ASCII_FOLD = str.maketrans(string.ascii_uppercase, string.ascii_lowercase)


class ScrubError(Exception):
    """A package the rules cannot be applied to cleanly (reported as an error, exit code 2)."""


def fail(msg: str) -> NoReturn:
    raise ScrubError(msg)


def part_key(name: str) -> str:
    """Comparison key of a part name. OPC compares part names ASCII-case-insensitively
    (ECMA-376 Part 2), and a zip entry name is the part name without its leading "/"."""
    return name.lstrip("/").translate(_ASCII_FOLD)


def index_entries(names: list[str], context: str) -> dict[str, str]:
    """Map each part key to the zip entry holding that part."""
    entries: dict[str, str] = {}
    for name in names:
        key = part_key(name)
        if key in entries:
            fail(f"{context}: zip entries {entries[key]!r} and {name!r} name the same part "
                 "(part names are case-insensitive)")
        entries[key] = name
    return entries


def resolve_target(rels_name: str, target: str) -> str:
    """The part an internal relationship target names, as a zip entry name (no leading "/").

    rels_name is the .rels part holding the relationship: <folder>/_rels/<source>.rels, or
    _rels/.rels for the package itself. A relative target resolves against the source part's
    folder, an absolute one ("/...") from the package root.
    """
    if not target.startswith("/"):
        source_folder = posixpath.dirname(posixpath.dirname(rels_name))
        target = posixpath.join("/", source_folder, target)
    return posixpath.normpath(target).lstrip("/")


def rels_part_for(part: str) -> str:
    """The relationships part of a part: <folder>/_rels/<name>.rels."""
    folder, _, name = part.rpartition("/")
    return posixpath.join(folder, "_rels", name + ".rels")


def splice(data: bytes, edits: list[tuple[int, int, bytes]]) -> bytes:
    """Replace each (start, end) byte range of data with its new bytes; ranges never overlap."""
    out, pos = [], 0
    for start, end, new in sorted(edits):
        out += [data[pos:start], new]
        pos = end
    out.append(data[pos:])
    return b"".join(out)


@dataclass
class Element:
    """An empty XML element located in a part's bytes, so it can be removed or edited in place."""
    start: int
    end: int
    attrs: dict[str, str]                # attribute values, XML-unescaped
    spans: dict[str, tuple[int, int]]    # byte range of each raw attribute value


_ATTRIBUTE = re.compile(rb"""([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')""")


def find_elements(data: bytes, local_name: str, context: str) -> list[Element]:
    """Every empty <local_name .../> element in data, under any namespace prefix.

    The byte-level scan is checked against a real XML parse, so XML serialised in a way it does
    not understand is refused instead of being silently skipped.
    """
    qname = rb"<(?:[\w.-]+:)?" + local_name.encode()
    pattern = re.compile(qname + rb"""((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*"""
                         rb"(?:/>|>\s*</(?:[\w.-]+:)?" + local_name.encode() + rb"\s*>)")
    found = []
    for m in pattern.finditer(data):
        attrs, spans = {}, {}
        for attr in _ATTRIBUTE.finditer(data, m.start(1), m.end(1)):
            group = 2 if attr.group(2) is not None else 3
            name = attr.group(1).decode()
            attrs[name] = html.unescape(attr.group(group).decode())
            spans[name] = attr.span(group)
        found.append(Element(m.start(), m.end(), attrs, spans))
    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        fail(f"{context}: not well-formed XML ({exc})")
    expected = sum(1 for el in root.iter() if el.tag.rpartition("}")[2] == local_name)
    if len(found) != expected:
        fail(f"{context}: read {len(found)} of its {expected} <{local_name}> elements; refusing "
             "to edit XML serialised in a way this script does not understand")
    return found


def is_external(rel: Element) -> bool:
    return rel.attrs.get("TargetMode") == "External"


def set_text(data: bytes, local_name: str, text: bytes, context: str,
             current: bytes = rb"[^<]+") -> tuple[bytes, str | None]:
    """Set the text of the <local_name> element (any namespace prefix) when it matches current;
    return the new data and the element's qualified name if it was changed."""
    changed: list[str] = []

    def replace(m: re.Match[bytes]) -> bytes:
        changed.append(m.group(1).decode())
        return b"<" + m.group(1) + b">" + text + b"</" + m.group(1) + b">"

    data = re.sub(rb"<((?:[\w.-]+:)?" + local_name.encode() + rb")>" + current + rb"</\1>",
                  replace, data)
    if len(changed) > 1:
        fail(f"{context}: <{changed[0]}> occurs {len(changed)} times, expected at most once")
    return data, changed[0] if changed else None


def scrub_core_props(data: bytes, context: str) -> tuple[bytes, list[str]]:
    actions: list[str] = []
    for local_name in ("creator", "lastModifiedBy"):
        data, name = set_text(data, local_name, b"", context)
        if name:
            actions.append(f"blanked <{name}>")
    data, name = set_text(data, "revision", b"1", context, current=rb"(?!1<)[^<]+")
    if name:
        actions.append(f"reset <{name}> to 1")
    return data, actions


def scrub_app_props(data: bytes, context: str) -> tuple[bytes, list[str]]:
    """Rule 7: blank the organisation and manager names in docProps/app.xml."""
    actions: list[str] = []
    for local_name in ("Company", "Manager"):
        data, name = set_text(data, local_name, b"", context)
        if name:
            actions.append(f"blanked <{name}>")
    return data, actions


FLAT_OPC_MARKER = "&lt;pkg:package"


def scrub_dataset_values(data: bytes, context: str) -> tuple[bytes, list[str]]:
    """Rule 6: replace flat-OPC-package sample values in a BC dataset part with the element name."""
    if b"NavWordReportXmlPart" not in data.replace(b"\x00", b""):
        return data, []
    if data[:2] == b"\xff\xfe":
        text = data.decode("utf-16")
        encode = lambda t: b"\xff\xfe" + t.encode("utf-16-le")  # noqa: E731
    else:
        bom = data[:3] == b"\xef\xbb\xbf"
        text = data.decode("utf-8-sig")
        encode = lambda t: (b"\xef\xbb\xbf" if bom else b"") + t.encode("utf-8")  # noqa: E731
    if FLAT_OPC_MARKER not in text:
        return data, []
    if encode(text) != data:
        fail(f"{context}: encoding does not round-trip byte-identically; refusing to edit")

    replaced: list[str] = []

    def substitute(m: re.Match[str]) -> str:
        name = m.group(1)
        replaced.append(name)
        return f"<{name}>{name}</{name}>"

    new_text = re.sub(
        r"<([A-Za-z_][\w.-]*)>([^<]*" + re.escape(FLAT_OPC_MARKER) + r"[^<]*)</\1>",
        substitute, text)
    if not replaced:
        fail(f"{context}: flat-OPC marker present but no enclosing leaf element matched")
    return encode(new_text), [f"replaced flat-OPC package sample value in <{n}>" for n in replaced]


def find_bibliography_sets(zf: zipfile.ZipFile, entries: dict[str, str],
                           context: str) -> list[list[str]]:
    """Every bibliography part, each as [item, its item rels, its itemProps] (those present)."""
    sets: list[list[str]] = []
    for name in zf.namelist():
        if not CUSTOM_XML_ITEM.fullmatch(name) or BIBLIOGRAPHY_NS not in zf.read(name)[:512]:
            continue
        members = [name]
        rels_name = entries.get(part_key(rels_part_for(name)))
        if rels_name is not None:
            members.append(rels_name)
            for rel in find_elements(zf.read(rels_name), "Relationship", f"{context}:{rels_name}"):
                target = resolve_target(rels_name, rel.attrs.get("Target", ""))
                if (is_external(rel) or not rel.attrs.get("Type", "").endswith("/customXmlProps")
                        or not CUSTOM_XML_PROPS.fullmatch(target)):
                    fail(f"{context}:{rels_name}: unexpected relationship {rel.attrs.get('Id')!r} "
                         f"({rel.attrs.get('Type')} -> {rel.attrs.get('Target')}); refusing to "
                         "guess whether it belongs with the bibliography part")
                props = entries.get(part_key(target))
                if props is not None and props not in members:
                    members.append(props)
        sets.append(members)
    return sets


def drop_relationships_to(data: bytes, rels_name: str, drop_keys: set[str],
                          context: str) -> tuple[bytes, list[str]]:
    """Remove every internal relationship whose target resolves to a dropped part; also return
    the key of the part each removed relationship pointed at."""
    removed = []
    for rel in find_elements(data, "Relationship", context):
        if is_external(rel):
            continue
        key = part_key(resolve_target(rels_name, rel.attrs.get("Target", "")))
        if key in drop_keys:
            removed.append((rel, key))
    return splice(data, [(rel.start, rel.end, b"") for rel, _ in removed]), [k for _, k in removed]


def drop_overrides(data: bytes, drop_keys: set[str], context: str) -> bytes:
    """Remove the [Content_Types].xml Override of every dropped part (a part has at most one)."""
    by_part: dict[str, list[Element]] = {}
    for override in find_elements(data, "Override", context):
        key = part_key(override.attrs.get("PartName", ""))
        if key in drop_keys:
            by_part.setdefault(key, []).append(override)
    for found in by_part.values():
        if len(found) > 1:
            fail(f"{context}: {len(found)} Overrides for {found[0].attrs['PartName']}; refusing "
                 "to edit an inconsistent package")
    return splice(data, [(o.start, o.end, b"") for found in by_part.values() for o in found])


def sanitize_attached_template(data: bytes, context: str) -> tuple[bytes, list[str]]:
    """Rewrite every external attachedTemplate target to the sanitized template directory."""
    actions: list[str] = []
    edits: list[tuple[int, int, bytes]] = []
    for rel in find_elements(data, "Relationship", context):
        if not (is_external(rel) and rel.attrs.get("Type", "").endswith("/attachedTemplate")):
            continue
        if "Target" not in rel.spans:
            fail(f"{context}: attachedTemplate relationship {rel.attrs.get('Id')!r} has no Target")
        start, end = rel.spans["Target"]
        target = data[start:end]
        filename = re.split(rb"[/\\]", target)[-1]
        sanitized = SANITIZED_TEMPLATE_DIR.encode() + filename
        if target != sanitized:
            edits.append((start, end, sanitized))
            actions.append(f"attachedTemplate target -> {sanitized.decode()}")
    return splice(data, edits), actions


def scrub_file(path: Path, check_only: bool) -> list[str]:
    """Scrub one .docx in place; return the list of actions taken (empty = already clean)."""
    actions: list[str] = []
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        infos = zf.infolist()
        entries = index_entries(names, path.name)
        content_types = entries.get(part_key(CONTENT_TYPES))
        if content_types is None:
            fail(f"{path.name}: no {CONTENT_TYPES}, so not an OPC package")

        edited: dict[str, bytes] = {}

        def current(name: str) -> bytes:
            return edited[name] if name in edited else zf.read(name)

        for part, scrub_props in ((CORE_PROPS, scrub_core_props), (APP_PROPS, scrub_app_props)):
            name = entries.get(part_key(part))
            if name is None:
                continue
            new_part, part_actions = scrub_props(zf.read(name), f"{path.name}:{name}")
            if part_actions:
                edited[name] = new_part
                actions += [f"{name}: {a}" for a in part_actions]

        bib_sets = find_bibliography_sets(zf, entries, path.name)
        # Parts dropped for what they are; each must be referenced from somewhere (see docstring).
        primary = [entries[key] for key in map(part_key, DROP_PARTS) if key in entries]
        primary += [members[0] for members in bib_sets]
        drop = set(primary).union(*bib_sets)

        for name in names:
            if CUSTOM_XML_ITEM.fullmatch(name) and name not in drop:
                new_part, part_actions = scrub_dataset_values(zf.read(name), f"{path.name}:{name}")
                if part_actions:
                    edited[name] = new_part
                    actions += [f"{name}: {a}" for a in part_actions]

        rels_parts = [name for name in names if RELS_PART.fullmatch(name) and name not in drop]
        if drop:
            actions += [f"dropped {p}" for p in sorted(drop)]
            drop_keys = {part_key(p) for p in drop}
            references = dict.fromkeys(map(part_key, primary), 0)
            for rels_name in rels_parts:
                new_rels, targets = drop_relationships_to(
                    current(rels_name), rels_name, drop_keys, f"{path.name}:{rels_name}")
                if targets:
                    edited[rels_name] = new_rels
                for key in targets:
                    if key in references:
                        references[key] += 1
            for name in primary:
                if not references[part_key(name)]:
                    fail(f"{path.name}: dropping {name}, but no relationship points at it; "
                         "refusing rather than risk leaving a reference this script cannot "
                         "resolve")
            ct = current(content_types)
            new_ct = drop_overrides(ct, drop_keys, f"{path.name}:{content_types}")
            if new_ct != ct:
                edited[content_types] = new_ct

        for rels_name in rels_parts:
            new_rels, at_actions = sanitize_attached_template(
                current(rels_name), f"{path.name}:{rels_name}")
            if at_actions:
                edited[rels_name] = new_rels
                actions += [f"{rels_name}: {a}" for a in at_actions]

        if not actions:
            return []
        if check_only:
            return actions

        tmp_fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=path.parent)
        os.close(tmp_fd)
        try:
            with zipfile.ZipFile(tmp_name, "w") as out:
                for info in infos:
                    if info.filename in drop:
                        continue
                    payload = edited.get(info.filename, zf.read(info.filename))
                    clone = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                    clone.compress_type = info.compress_type
                    clone.external_attr = info.external_attr
                    clone.internal_attr = info.internal_attr
                    clone.create_system = info.create_system
                    out.writestr(clone, payload)
        except BaseException:
            Path(tmp_name).unlink(missing_ok=True)
            raise
    # Only once the source zip is closed: on Windows a rename cannot replace a file that is still
    # open, and shutil.move then silently degraded to copying over it, which a crash could leave
    # half-written. os.replace within one folder is a single atomic rename.
    try:
        os.replace(tmp_name, path)
    except BaseException:
        Path(tmp_name).unlink(missing_ok=True)
        raise
    return actions


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        prog="scrub_corpus_metadata.py", allow_abbrev=False,
        description="Scrub personal metadata from corpus .docx files (rules: see the docstring).")
    parser.add_argument("--check", action="store_true",
                        help="report what would change and exit 1 if anything would; write nothing")
    parser.add_argument("paths", nargs="*", type=Path,
                        help="the .docx files to scrub (default: tests/corpus/*.docx)")
    args = parser.parse_args(argv)  # a bad command line exits 2 with a usage line
    targets = args.paths
    if not targets:
        repo_root = Path(__file__).resolve().parent.parent
        targets = sorted((repo_root / "tests" / "corpus").glob("*.docx"))
    if not targets:
        print("scrub_corpus_metadata: ERROR: no corpus files found", file=sys.stderr)
        return 2

    dirty = failed = False
    for path in targets:
        try:
            if not path.is_file():
                fail(f"not a file: {path}")
            actions = scrub_file(path, args.check)
        except ScrubError as exc:
            print(f"scrub_corpus_metadata: ERROR: {exc}", file=sys.stderr)
            failed = True
            continue
        except Exception as exc:  # e.g. not a zip at all: an error (2), never "dirty" (1)
            print(f"scrub_corpus_metadata: ERROR: {path.name}: {type(exc).__name__}: {exc}",
                  file=sys.stderr)
            failed = True
            continue
        if actions:
            dirty = True
            verb = "would scrub" if args.check else "scrubbed"
            print(f"{verb} {path.name}:")
            for a in actions:
                print(f"  - {a}")
        else:
            print(f"clean   {path.name}")
    if failed:
        return 2
    return 1 if (args.check and dirty) else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
