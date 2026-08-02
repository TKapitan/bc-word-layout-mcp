"""Re-runs the mock preview on the REAL data BC used, so mock and BC renders are comparable.

The sandbox pack's mock previews use generated sample data, so a side-by-side against a BC render
mixes two different questions: "does the layout render the same" and "is the data the same". This
script removes the second one. For every pack item that has a BC dataset export in its bc-output/, it:

    1. converts the export into the shape preview_layout's dataOverridesPath accepts,
    2. re-runs preview_layout on the SAME layout with that data,
    3. renders the result to PNGs beside the BC render's PNGs.

    python tools/e2e/bc_compare.py            # every item with a bc-output/*.xml
    python tools/e2e/bc_compare.py p01 p04    # specific items

Output per item, under preview-output/sandbox-pack/<item>/:

    bc-output/overrides.xml        the converted dataset (what the mock was fed)
    bc-data/mock-preview.pdf       the mock rendered from BC's own data
    bc-data/pages/page-N.png       its pages, to set against bc-output/pages/page-N.png

WHY A CONVERTER IS NEEDED (a real gap, not an accident of this script): what Business Central's UI
exports via *Send to > XML* is a generic

    <ReportDataSet name id language wordMergeDataItem>
      <Labels><Label name="X">text</Label></Labels>
      <DataItems><DataItem name="Header"><Columns><Column name="Y">value</Column></Columns>
                                         <DataItems>...nested...</DataItems></DataItem></DataItems>

whereas dataOverridesPath requires the layout's own data-store part shape - a NavWordReportXmlPart
root in the report's urn:microsoft-dynamics-nav namespace, with each column as an ELEMENT named after the
column. Both carry the same information; only the encoding differs. Until the server accepts the
export shape directly (GitHub issue #4), this conversion is the bridge.
"""

from __future__ import annotations

import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scenarios import McpServer  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
PACK_ROOT = os.path.join(REPO, "preview-output", "sandbox-pack")
RENDER_PAGES = 3


def layout_namespace(docx_path: str) -> str:
    """The BC dataset part's namespace - the namespace the converted dataset must live in."""
    with zipfile.ZipFile(docx_path) as z:
        for name in z.namelist():
            if not (name.startswith("customXml/item") and name.endswith(".xml")):
                continue
            raw = z.read(name)
            text = raw.decode("utf-16") if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else \
                raw.decode("utf-8", "ignore")
            m = re.search(r'xmlns="(urn:microsoft-dynamics-nav/reports/[^"]+)"', text)
            if m:
                return m.group(1)
    raise SystemExit(f"no BC dataset part found in {docx_path}")


def convert(export_path: str, ns: str) -> ET.ElementTree:
    """BC's ReportDataSet export -> the NavWordReportXmlPart shape, in the layout's namespace."""
    src = ET.parse(export_path).getroot()
    ET.register_namespace("", ns)
    root = ET.Element(f"{{{ns}}}NavWordReportXmlPart")

    def emit_columns(container: ET.Element, target: ET.Element):
        cols = container.find("Columns")
        for col in (cols if cols is not None else []):
            name = col.get("name")
            if not name:
                continue
            el = ET.SubElement(target, f"{{{ns}}}{name}")
            el.text = col.text or ""

    def emit_items(container: ET.Element, target: ET.Element):
        items = container.find("DataItems")
        for item in (items if items is not None else []):
            name = item.get("name")
            if not name:
                continue
            # One element per DataItem occurrence: BC emits a sibling per row, which is exactly what
            # the merge engine expands a repeater over.
            el = ET.SubElement(target, f"{{{ns}}}{name}")
            emit_columns(item, el)
            emit_items(item, el)

    # Report metadata, when present, mirrors the layout's own /BCReportInformation subtree.
    def emit_verbatim(node: ET.Element, target: ET.Element):
        for child in node:
            el = ET.SubElement(target, f"{{{ns}}}{child.tag}")
            el.text = child.text or ""
            emit_verbatim(child, el)

    info = src.find("BCReportInformation")
    if info is not None:
        emit_verbatim(info, ET.SubElement(root, f"{{{ns}}}BCReportInformation"))

    labels = src.find("Labels")
    if labels is not None:
        target = ET.SubElement(root, f"{{{ns}}}Labels")
        for label in labels:
            name = label.get("name")
            if name:
                el = ET.SubElement(target, f"{{{ns}}}{name}")
                el.text = label.text or ""

    emit_items(src, root)
    return ET.ElementTree(root)


def main() -> int:
    wanted = sys.argv[1:]
    items = sorted(d for d in os.listdir(PACK_ROOT)
                   if os.path.isdir(os.path.join(PACK_ROOT, d))
                   and (not wanted or any(d.startswith(w) for w in wanted)))
    server = McpServer()
    rows = []
    try:
        for item in items:
            item_dir = os.path.join(PACK_ROOT, item)
            bc_dir = os.path.join(item_dir, "bc-output")
            exports = [f for f in os.listdir(bc_dir) if f.endswith(".xml") and f != "overrides.xml"] \
                if os.path.isdir(bc_dir) else []
            layouts = [f for f in os.listdir(item_dir) if f.endswith(".docx") and "merged" not in f]
            if not exports or not layouts:
                rows.append((item, "skipped", "no BC dataset export" if not exports else "no layout"))
                continue

            layout = os.path.join(item_dir, layouts[0])
            export = os.path.join(bc_dir, sorted(exports)[-1])   # newest round wins
            ns = layout_namespace(layout)
            overrides = os.path.join(bc_dir, "overrides.xml")
            convert(export, ns).write(overrides, encoding="utf-8", xml_declaration=True)

            out_dir = os.path.join(item_dir, "bc-data")
            os.makedirs(out_dir, exist_ok=True)
            env = server.tool("preview_layout", {"layoutPath": layout, "dataOverridesPath": overrides,
                                                 "outputDir": out_dir})
            if not env.get("ok"):
                rows.append((item, "FAILED", str(env.get("error"))[:120]))
                continue
            data = env["data"]
            stats = data.get("stats") or {}
            pdf = data.get("pdfPath")
            if pdf:
                target = os.path.join(out_dir, "mock-preview.pdf")
                if os.path.abspath(pdf) != os.path.abspath(target):
                    os.replace(pdf, target)
                server.tool("render_preview_pages", {"pdfPath": target, "maxPages": RENDER_PAGES,
                                                     "dpi": 110},
                            image_dir=os.path.join(out_dir, "pages"))
            rows.append((item, "ok" if pdf else "no pdf",
                         f"fields={stats.get('fieldsFilled')} rows={stats.get('rowsGenerated')} "
                         f"unresolved={stats.get('unresolved')} warnings={len(data.get('warnings') or [])}"))
    finally:
        server.close()

    width = max(len(r[0]) for r in rows) if rows else 10
    for name, status, detail in rows:
        print(f"{name:<{width}}  {status:<8} {detail}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
