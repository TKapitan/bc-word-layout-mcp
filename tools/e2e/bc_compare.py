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

Exit status: 0 when every compared item rendered (an item with no BC export or no layout yet is
"skipped", not failed); 1 when any item FAILED - a failed preview, a conversion that produced no PDF, a
failed or empty render, or an error; 2 when there is nothing to run against: no server build, no pack,
or no item matching the arguments.

The *Send to > XML* export (a ReportDataSet document) is handed to preview_layout AS-IS: since
GitHub issue #4 the server accepts it directly - converting it to the layout's data-store part
shape internally, decimalformatter columns formatted with the export's formatRegion culture - so
the converter bridge that used to live here (and never applied decimalformatter) is gone.
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scenarios import HOST, McpServer, describe_error  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
PACK_ROOT = os.path.join(REPO, "preview-output", "sandbox-pack")
RENDER_PAGES = 3
FAILED = "FAILED"


def exit_code(rows: list[tuple[str, str, str]]) -> int:
    """The script's exit status for its (item, status, detail) rows: 1 when any row FAILED, else 0.

    An "ok" row rendered; a "skipped" row had nothing to compare yet (no BC export or no layout), which
    is not a failure. The script used to return 0 whatever its rows said (GitHub issue #112).
    """
    return 1 if any(status == FAILED for _, status, _ in rows) else 0


def compare_item(server: McpServer, item_dir: str) -> tuple[str, str]:
    """Previews one pack item on its newest BC export and renders the result: (status, detail)."""
    bc_dir = os.path.join(item_dir, "bc-output")
    # overrides.xml is the converter bridge's old output; skip it in packs that predate #4.
    exports = [f for f in os.listdir(bc_dir) if f.endswith(".xml") and f != "overrides.xml"] \
        if os.path.isdir(bc_dir) else []
    layouts = [f for f in os.listdir(item_dir) if f.endswith(".docx") and "merged" not in f]
    if not exports or not layouts:
        return "skipped", "no BC dataset export" if not exports else "no layout"

    layout = os.path.join(item_dir, layouts[0])
    export = os.path.join(bc_dir, sorted(exports)[-1])   # newest round wins

    out_dir = os.path.join(item_dir, "bc-data")
    os.makedirs(out_dir, exist_ok=True)
    env = server.tool("preview_layout", {"layoutPath": layout, "dataOverridesPath": export,
                                         "outputDir": out_dir})
    if not env.get("ok"):
        return FAILED, f"preview_layout failed: {describe_error(env)}"
    data = env.get("data") or {}
    stats = data.get("stats") or {}
    summary = (f"fields={stats.get('fieldsFilled')} rows={stats.get('rowsGenerated')} "
               f"unresolved={stats.get('unresolved')} warnings={len(data.get('warnings') or [])}")
    pdf = data.get("pdfPath")
    if not data.get("conversionOk") or not pdf:
        return FAILED, (f"no PDF (converter {data.get('converterUsed')}): "
                        f"{data.get('conversionError') or 'no conversion error reported'}; {summary}")
    target = os.path.join(out_dir, "mock-preview.pdf")
    if os.path.abspath(pdf) != os.path.abspath(target):
        os.replace(pdf, target)
    render = server.tool("render_preview_pages", {"pdfPath": target, "maxPages": RENDER_PAGES,
                                                  "dpi": 110},
                         image_dir=os.path.join(out_dir, "pages"))
    if not render.get("ok"):
        return FAILED, f"render_preview_pages failed: {describe_error(render)}; {summary}"
    pages = render.get("data") or {}
    if not pages.get("pagesRendered"):
        return FAILED, f"render_preview_pages rendered no page; {summary}"
    return "ok", f"{summary} pages={pages['pagesRendered']}/{pages.get('pageCount')}"


def main(argv: list[str] | None = None) -> int:
    wanted = sys.argv[1:] if argv is None else argv
    if not os.path.exists(HOST):
        print(f"server not built: {HOST}\nrun: dotnet build -c Release", file=sys.stderr)
        return 2
    if not os.path.isdir(PACK_ROOT):
        print(f"no sandbox pack at {PACK_ROOT}\nbuild it first: python tools/e2e/sandbox_pack.py",
              file=sys.stderr)
        return 2
    items = sorted(d for d in os.listdir(PACK_ROOT)
                   if os.path.isdir(os.path.join(PACK_ROOT, d))
                   and (not wanted or any(d.startswith(w) for w in wanted)))
    if not items:
        print(f"no pack item matches {wanted}" if wanted else f"the pack at {PACK_ROOT} has no items",
              file=sys.stderr)
        return 2

    try:
        server = McpServer()
    except (RuntimeError, OSError) as exc:  # ServerExited, TimeoutError, a refused handshake
        print(f"could not start the server: {exc}", file=sys.stderr)
        return 1
    rows = []
    try:
        for item in items:
            try:
                status, detail = compare_item(server, os.path.join(PACK_ROOT, item))
            except Exception as exc:  # noqa: BLE001 - one item's error is that item's FAILED row
                status, detail = FAILED, f"{type(exc).__name__}: {exc}"
            rows.append((item, status, detail))
    finally:
        server.close()

    width = max(len(r[0]) for r in rows)
    for name, status, detail in rows:
        print(f"{name:<{width}}  {status:<8} {detail}")
    return exit_code(rows)


if __name__ == "__main__":
    sys.exit(main())
