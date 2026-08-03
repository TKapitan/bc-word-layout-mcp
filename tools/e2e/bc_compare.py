"""Re-runs the mock preview on the REAL data BC used, so mock and BC renders are comparable.

The sandbox pack's mock previews use generated sample data, so a side-by-side against a BC render
mixes two different questions: "does the layout render the same" and "is the data the same". This
script removes the second one. For every pack item that has a BC dataset export in its bc-output/, it:

    1. re-runs preview_layout on the SAME layout with that export as dataOverridesPath,
    2. renders the result to PNGs beside the BC render's PNGs.

    python tools/e2e/bc_compare.py            # every item with a bc-output/*.xml
    python tools/e2e/bc_compare.py p01 p04    # specific items

Output per item, under preview-output/sandbox-pack/<item>/:

    bc-data/mock-preview.pdf       the mock rendered from BC's own data
    bc-data/pages/page-N.png       its pages, to set against bc-output/pages/page-N.png

The *Send to > XML* export (a ReportDataSet document) is handed to preview_layout AS-IS: since
GitHub issue #4 the server accepts it directly - converting it to the layout's data-store part
shape internally, decimalformatter columns formatted with the export's formatRegion culture - so
the converter bridge that used to live here (and never applied decimalformatter) is gone.
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scenarios import McpServer  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
PACK_ROOT = os.path.join(REPO, "preview-output", "sandbox-pack")
RENDER_PAGES = 3


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
            # overrides.xml is the converter bridge's old output; skip it in packs that predate #4.
            exports = [f for f in os.listdir(bc_dir) if f.endswith(".xml") and f != "overrides.xml"] \
                if os.path.isdir(bc_dir) else []
            layouts = [f for f in os.listdir(item_dir) if f.endswith(".docx") and "merged" not in f]
            if not exports or not layouts:
                rows.append((item, "skipped", "no BC dataset export" if not exports else "no layout"))
                continue

            layout = os.path.join(item_dir, layouts[0])
            export = os.path.join(bc_dir, sorted(exports)[-1])   # newest round wins

            out_dir = os.path.join(item_dir, "bc-data")
            os.makedirs(out_dir, exist_ok=True)
            env = server.tool("preview_layout", {"layoutPath": layout, "dataOverridesPath": export,
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
