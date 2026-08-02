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
  4. bibliography customXml - DROPPED where present. An empty <b:Sources> part Word inserts
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

Everything else - the BC dataset customXml part (UTF-16, byte-sensitive), the document body,
headers/footers, media, styles, every other relationship - is copied through byte-identical.
Zip entry order and per-entry timestamps are preserved.

Idempotent: a second run reports every file unchanged. Run it whenever a new capture joins
the corpus. Only top-level tests/corpus/*.docx are touched (subfolders are private material).

Usage:  python tools/scrub_corpus_metadata.py [--check] [paths...]
        --check  report what would change and exit 1 if anything would; write nothing.
"""

from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

BIBLIOGRAPHY_NS = b"http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
SANITIZED_TEMPLATE_DIR = "file:///C:/Templates/"

# Parts dropped outright wherever they appear, with the package plumbing that references them.
DROP_PARTS = ("docProps/custom.xml", "docMetadata/LabelInfo.xml")


def fail(msg: str) -> None:
    raise SystemExit(f"scrub_corpus_metadata: ERROR: {msg}")


def sub_expect(pattern: bytes, replacement: bytes, data: bytes, *, at_most: int | None = None,
               context: str = "") -> tuple[bytes, int]:
    """re.sub that reports how many replacements happened and enforces an optional ceiling."""
    new, count = re.subn(pattern, replacement, data)
    if at_most is not None and count > at_most:
        fail(f"{context}: pattern {pattern!r} matched {count} times, expected at most {at_most}")
    return new, count


def scrub_core_props(data: bytes, context: str) -> tuple[bytes, list[str]]:
    actions: list[str] = []
    for tag in (b"dc:creator", b"cp:lastModifiedBy"):
        data, n = sub_expect(
            rb"<" + tag + rb">[^<]+</" + tag + rb">",
            b"<" + tag + b"></" + tag + b">",
            data, at_most=1, context=context)
        if n:
            actions.append(f"blanked <{tag.decode()}>")
    data, n = sub_expect(
        rb"<cp:revision>(?!1<)[^<]+</cp:revision>",
        b"<cp:revision>1</cp:revision>",
        data, at_most=1, context=context)
    if n:
        actions.append("reset <cp:revision> to 1")
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


def find_bibliography_set(zf: zipfile.ZipFile) -> tuple[list[str], str | None]:
    """Return ([item, item rels, itemProps] of the bibliography part, item index) or ([], None)."""
    for name in zf.namelist():
        m = re.fullmatch(r"customXml/item(\d+)\.xml", name)
        if not m:
            continue
        head = zf.read(name)[:512]
        if BIBLIOGRAPHY_NS not in head:
            continue
        idx = m.group(1)
        members = [name]
        for extra in (f"customXml/_rels/item{idx}.xml.rels", f"customXml/itemProps{idx}.xml"):
            if extra in zf.namelist():
                members.append(extra)
        return members, idx
    return [], None


def drop_relationship_by_target(data: bytes, target: bytes, context: str) -> tuple[bytes, int]:
    return sub_expect(
        rb'<Relationship [^>]*Target="' + re.escape(target) + rb'"[^>]*/>',
        b"", data, at_most=1, context=context)


def drop_content_type_override(data: bytes, part_name: bytes, context: str) -> tuple[bytes, int]:
    return sub_expect(
        rb'<Override PartName="' + re.escape(part_name) + rb'"[^>]*/>',
        b"", data, at_most=1, context=context)


def sanitize_attached_template(data: bytes, context: str) -> tuple[bytes, list[str]]:
    """Rewrite every external attachedTemplate target to the sanitized template directory."""
    actions: list[str] = []

    def rewrite(match: re.Match[bytes]) -> bytes:
        rel = match.group(0)
        tgt = re.search(rb'Target="([^"]+)"', rel)
        assert tgt is not None
        target = tgt.group(1)
        filename = re.split(rb"[/\\]", target)[-1]
        sanitized = SANITIZED_TEMPLATE_DIR.encode() + filename
        if target == sanitized:
            return rel
        actions.append(f'attachedTemplate target -> {sanitized.decode()}')
        return rel.replace(b'Target="' + target + b'"', b'Target="' + sanitized + b'"')

    new = re.sub(
        rb'<Relationship [^>]*Type="[^"]*/attachedTemplate"[^>]*TargetMode="External"[^>]*/>',
        rewrite, data)
    return new, actions


def scrub_file(path: Path, check_only: bool) -> list[str]:
    """Scrub one .docx in place; return the list of actions taken (empty = already clean)."""
    actions: list[str] = []
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        infos = zf.infolist()
        bib_members, bib_idx = find_bibliography_set(zf)
        drop = {p for p in DROP_PARTS if p in names} | set(bib_members)

        edited: dict[str, bytes] = {}

        core = zf.read("docProps/core.xml")
        new_core, core_actions = scrub_core_props(core, f"{path.name}:docProps/core.xml")
        if core_actions:
            edited["docProps/core.xml"] = new_core
            actions += [f"docProps/core.xml: {a}" for a in core_actions]

        for name in names:
            if re.fullmatch(r"customXml/item\d+\.xml", name) and name not in drop:
                new_part, part_actions = scrub_dataset_values(zf.read(name), f"{path.name}:{name}")
                if part_actions:
                    edited[name] = new_part
                    actions += [f"{name}: {a}" for a in part_actions]

        if drop:
            actions += [f"dropped {p}" for p in sorted(drop)]

            root_rels = zf.read("_rels/.rels")
            for part in DROP_PARTS:
                if part in drop:
                    root_rels, n = drop_relationship_by_target(
                        root_rels, part.encode(), f"{path.name}:_rels/.rels")
                    if not n:
                        fail(f"{path.name}: dropping {part} but no rel in _rels/.rels points at it")
            if root_rels != zf.read("_rels/.rels"):
                edited["_rels/.rels"] = root_rels

            ct = zf.read("[Content_Types].xml")
            for part in DROP_PARTS:
                if part in drop:
                    ct, _ = drop_content_type_override(
                        ct, b"/" + part.encode(), f"{path.name}:[Content_Types].xml")
            if bib_idx is not None:
                ct, _ = drop_content_type_override(
                    ct, f"/customXml/itemProps{bib_idx}.xml".encode(),
                    f"{path.name}:[Content_Types].xml")
            if ct != zf.read("[Content_Types].xml"):
                edited["[Content_Types].xml"] = ct

            if bib_idx is not None:
                doc_rels = zf.read("word/_rels/document.xml.rels")
                doc_rels, n = drop_relationship_by_target(
                    doc_rels, f"../customXml/item{bib_idx}.xml".encode(),
                    f"{path.name}:word/_rels/document.xml.rels")
                if not n:
                    fail(f"{path.name}: bibliography item{bib_idx} has no document.xml.rels entry")
                edited["word/_rels/document.xml.rels"] = doc_rels

        for rels_name in names:
            if not rels_name.endswith(".rels") or rels_name in drop:
                continue
            base = edited.get(rels_name, zf.read(rels_name))
            new, at_actions = sanitize_attached_template(base, f"{path.name}:{rels_name}")
            if at_actions:
                edited[rels_name] = new
                actions += [f"{rels_name}: {a}" for a in at_actions]

        if not actions:
            return []
        if check_only:
            return actions

        tmp_fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=path.parent)
        import os
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
        shutil.move(tmp_name, path)
    return actions


def main(argv: list[str]) -> int:
    args = [a for a in argv if not a.startswith("--")]
    check_only = "--check" in argv
    if args:
        targets = [Path(a) for a in args]
    else:
        repo_root = Path(__file__).resolve().parent.parent
        targets = sorted((repo_root / "tests" / "corpus").glob("*.docx"))
    if not targets:
        fail("no corpus files found")

    dirty = False
    for path in targets:
        if not path.is_file():
            fail(f"not a file: {path}")
        actions = scrub_file(path, check_only)
        if actions:
            dirty = True
            verb = "would scrub" if check_only else "scrubbed"
            print(f"{verb} {path.name}:")
            for a in actions:
                print(f"  - {a}")
        else:
            print(f"clean   {path.name}")
    return 1 if (check_only and dirty) else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
