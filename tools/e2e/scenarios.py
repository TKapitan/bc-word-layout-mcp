"""End-to-end edit-scenario harness for the bc-word-layout MCP server.

Runs realistic multi-tool editing scenarios ("remove column X from the lines", "add field Y to the
header", ...) against FRESH COPIES of the corpus layouts through the real stdio MCP server, then
validates + previews + renders the result so a human (or vision-capable agent) can review the
before/after PNGs.

    python tools/e2e/scenarios.py               # run everything not yet run (skips cached passes)
    python tools/e2e/scenarios.py --list        # list scenarios and their cached status
    python tools/e2e/scenarios.py s04 s07       # run specific scenarios (prefix match)
    python tools/e2e/scenarios.py --force s04   # rerun even if cached

CACHING: each scenario's output dir carries a done.json with a hash of the scenario's own source
code; a scenario is skipped when it previously PASSED and its code is unchanged, so adding a new
scenario never reruns the existing ones. --force overrides.

OUTPUT (gitignored, human-review artifacts):
    preview-output/e2e-edits/_baseline/<layout>/   validate findings + page PNGs of the pristine layout
    preview-output/e2e-edits/<scenario>/           work.docx, report.json, done.json, after/page-N.png

The server is the Release build; build it first (dotnet build -c Release). Keeps stdin open until
responses arrive (stdio EOF race - see the pilot-harness notes).
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import inspect
import json
import os
import shutil
import subprocess
import sys
import threading
import time
import traceback

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
HOST = os.path.join(REPO, "src", "BcWordLayout.McpHost", "bin", "Release", "net10.0", "BcWordLayout.McpHost.dll")
CORPUS_DIR = os.path.join(REPO, "tests", "corpus")
OUT_ROOT = os.path.join(REPO, "preview-output", "e2e-edits")
BASELINE_ROOT = os.path.join(OUT_ROOT, "_baseline")
RENDER_PAGES = 2


# ---------------------------------------------------------------- MCP driver

class McpServer:
    """One persistent stdio MCP server shared by every scenario in a run."""

    def __init__(self, env: dict | None = None):
        """env: extra environment variables for the server process (e.g. the BCWL_LABEL_* knobs,
        which are read once at startup - a different value needs a different server)."""
        self.proc = subprocess.Popen(
            ["dotnet", HOST], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL, cwd=REPO,
            env={**os.environ, **env} if env else None)
        self.responses: dict[int, dict] = {}
        self.event = threading.Event()
        self.next_id = 0
        threading.Thread(target=self._reader, daemon=True).start()
        self._call("initialize", {"protocolVersion": "2024-11-05", "capabilities": {},
                                  "clientInfo": {"name": "e2e-scenarios", "version": "1"}})
        self._notify("notifications/initialized")

    def _reader(self):
        for line in self.proc.stdout:
            try:
                msg = json.loads(line.strip() or "{}")
            except json.JSONDecodeError:
                continue
            if "id" in msg:
                self.responses[msg["id"]] = msg
                self.event.set()

    def _send(self, payload: dict):
        self.proc.stdin.write((json.dumps(payload) + "\n").encode())
        self.proc.stdin.flush()

    def _notify(self, method: str):
        self._send({"jsonrpc": "2.0", "method": method})

    def _call(self, method: str, params: dict, timeout: int = 300) -> dict:
        self.next_id += 1
        mid = self.next_id
        self._send({"jsonrpc": "2.0", "id": mid, "method": method, "params": params})
        deadline = time.monotonic() + timeout
        while mid not in self.responses:
            self.event.clear()
            if not self.event.wait(min(10, max(0.1, deadline - time.monotonic()))) \
                    and time.monotonic() > deadline:
                raise TimeoutError(f"MCP call {method} timed out after {timeout}s")
        return self.responses.pop(mid)

    def tool(self, name: str, args: dict, image_dir: str | None = None) -> dict:
        """Call a tool; returns the parsed envelope. Saves image blocks as page-N.png if image_dir."""
        resp = self._call("tools/call", {"name": name, "arguments": args})
        result = resp.get("result")
        if result is None:
            return {"ok": False, "error": {"code": "jsonrpc_error", "message": json.dumps(resp.get("error"))}}
        envelope, saved = None, 0
        for block in result.get("content", []):
            if block.get("type") == "text" and envelope is None:
                try:
                    envelope = json.loads(block["text"])
                except json.JSONDecodeError:
                    envelope = {"ok": False, "error": {"code": "non_envelope_text", "message": block["text"]}}
            elif block.get("type") == "image" and image_dir:
                saved += 1
                os.makedirs(image_dir, exist_ok=True)
                with open(os.path.join(image_dir, f"page-{saved}.png"), "wb") as f:
                    f.write(base64.b64decode(block["data"]))
        return envelope if envelope is not None else {"ok": False, "error": {"code": "no_text_block", "message": "?"}}

    def close(self):
        try:
            self.proc.stdin.close()
            self.proc.wait(timeout=10)
        except Exception:
            self.proc.kill()


# ---------------------------------------------------------------- scenario context

class StepFailure(Exception):
    pass


class Ctx:
    """What a scenario function gets: tool sugar, layout-inspection helpers, assertions."""

    def __init__(self, server: McpServer, work: str | None, out_dir: str):
        self.server = server
        self.work = work            # absolute path of the scenario's working .docx (None until created)
        self.out_dir = out_dir
        self.steps: list[dict] = []
        self.checks: list[dict] = []

    # -- tool calls --------------------------------------------------------
    def tool(self, name: str, args: dict, expect_ok: bool = True) -> dict:
        env = self.server.tool(name, args)
        self.steps.append({"tool": name, "args": args, "ok": env.get("ok"),
                           "summary": (env.get("data") or {}).get("summary") if isinstance(env.get("data"), dict) else None,
                           "error": env.get("error")})
        if expect_ok and not env.get("ok"):
            raise StepFailure(f"{name} failed: {json.dumps(env.get('error'))[:300]}")
        if not expect_ok and env.get("ok"):
            raise StepFailure(f"{name} unexpectedly SUCCEEDED (a rejection was the expected behavior)")
        return env

    def edit(self, name: str, expect_ok: bool = True, **kwargs) -> dict:
        return self.tool(name, {"layoutPath": self.work, **kwargs}, expect_ok=expect_ok)

    # -- layout inspection -------------------------------------------------
    def info(self) -> dict:
        return self.tool("get_layout_info", {"layoutPath": self.work})["data"]

    def fields_root(self) -> dict:
        return self.tool("list_dataset_fields", {"source": self.work})["data"]["root"]

    def validate(self) -> dict:
        return self.tool("validate_layout", {"layoutPath": self.work, "level": "full"})["data"]

    # -- discovery helpers ---------------------------------------------------
    @staticmethod
    def grid_total(table: dict) -> int:
        return sum(table["gridColumnWidths"])

    @staticmethod
    def find_data_item(node: dict, name: str) -> dict | None:
        if node["name"] == name:
            return node
        for child in node.get("children", []):
            found = Ctx.find_data_item(child, name)
            if found:
                return found
        return None

    @staticmethod
    def unbound_field_with_label(item: dict) -> tuple[str, str]:
        """First unbound non-label column of the data item that has a matching *_Lbl/*Lbl sibling.
        Returns (fieldPath, labelPath)."""
        labels = {c["name"]: c["path"] for c in item.get("columns", []) if c["isLabel"]}
        for c in item.get("columns", []):
            if c["isLabel"] or c.get("bound"):
                continue
            for candidate in (c["name"] + "_Lbl", c["name"] + "Lbl"):
                if candidate in labels:
                    return c["path"], labels[candidate]
        raise StepFailure(f"no unbound field with a label found under {item['name']}")

    # -- assertions ----------------------------------------------------------
    def expect(self, condition: bool, label: str):
        self.checks.append({"label": label, "passed": bool(condition)})
        if not condition:
            raise StepFailure(f"check failed: {label}")

    def note(self, label: str):
        self.checks.append({"label": label, "passed": True, "note": True})


# ---------------------------------------------------------------- registry / runner

SCENARIOS: list[dict] = []


def scenario(name: str, layout: str | None, desc: str, allow_new_findings: bool = False,
             render_pages: int = RENDER_PAGES):
    """layout: corpus file the scenario edits a copy of; None = the scenario creates its own work.docx."""
    def register(fn):
        SCENARIOS.append({"name": name, "layout": layout, "desc": desc, "fn": fn,
                          "allow_new_findings": allow_new_findings, "render_pages": render_pages})
        return fn
    return register


def finding_signature(finding: dict) -> tuple:
    return (finding.get("severity"), finding.get("check"), finding.get("location"))


def scenario_hash(entry: dict) -> str:
    """Cache key for a scenario's recorded pass: its own source AND the server build it ran against.

    Source alone is not enough. A change in the SERVER can flip a scenario's outcome while its source is
    untouched, and the cached pass then keeps the stale expectation alive indefinitely - s02 pinned interior
    insert_column as a rejection for a full release after it had been implemented, because nothing invalidated
    its 'PASSED' record.
    """
    try:
        st = os.stat(HOST)
        build = f"{st.st_mtime_ns}:{st.st_size}"
    except OSError:
        build = "no-build"
    return hashlib.sha256((inspect.getsource(entry["fn"]) + "|" + build).encode()).hexdigest()[:16]


def baseline_stamp(layout: str) -> str:
    """Identity of the inputs a cached baseline depends on: the server build and the corpus file itself.

    Without this the cache silently outlives the thing it describes. A change to LayoutValidator's checks
    makes every cached "before" set incomplete, so each scenario's no-NEW-findings check reports the new
    findings as if the EDIT introduced them - three scenarios failed exactly that way after the 2026-08-01
    binding-namespace check landed, and `--force` does not help because it only reruns scenarios. Worse than
    the false failures: a stale baseline can also hide a real new finding whose signature it happens to
    contain already.
    """
    parts = []
    for path in (HOST, os.path.join(CORPUS_DIR, layout)):
        try:
            st = os.stat(path)
            parts.append(f"{os.path.basename(path)}:{st.st_mtime_ns}:{st.st_size}")
        except OSError:
            parts.append(f"{os.path.basename(path)}:missing")
    return hashlib.sha256("|".join(parts).encode()).hexdigest()[:16]


def ensure_baseline(server: McpServer, layout: str) -> dict:
    """Validate + preview + render the PRISTINE corpus layout once per layout (shared 'before')."""
    base_dir = os.path.join(BASELINE_ROOT, os.path.splitext(layout)[0])
    report_path = os.path.join(base_dir, "report.json")
    stamp = baseline_stamp(layout)
    if os.path.exists(report_path):
        with open(report_path) as f:
            cached = json.load(f)
        if cached.get("stamp") == stamp:
            return cached
        print(f"      baseline for {layout} is stale (server or corpus file changed) - recapturing")
    os.makedirs(base_dir, exist_ok=True)
    source = os.path.join(CORPUS_DIR, layout)
    validation = server.tool("validate_layout", {"layoutPath": source, "level": "full"})["data"]
    preview = server.tool("preview_layout", {"layoutPath": source, "outputDir": base_dir})
    report = {"layout": layout, "stamp": stamp, "validation": validation,
              "previewStats": preview["data"]["stats"] if preview.get("ok") else None}
    if preview.get("ok") and preview["data"].get("pdfPath"):
        server.tool("render_preview_pages",
                    {"pdfPath": preview["data"]["pdfPath"], "maxPages": RENDER_PAGES}, image_dir=base_dir)
    with open(report_path, "w") as f:
        json.dump(report, f, indent=2)
    return report


def run_scenario(server: McpServer, entry: dict) -> dict:
    name = entry["name"]
    out_dir = os.path.join(OUT_ROOT, name)
    os.makedirs(out_dir, exist_ok=True)
    work = os.path.join(out_dir, "work.docx")

    baseline = None
    if entry["layout"]:
        baseline = ensure_baseline(server, entry["layout"])
        shutil.copyfile(os.path.join(CORPUS_DIR, entry["layout"]), work)

    ctx = Ctx(server, work if entry["layout"] else None, out_dir)
    report = {"scenario": name, "desc": entry["desc"], "layout": entry["layout"], "passed": False}
    try:
        entry["fn"](ctx)

        # ---- common post-verification: validate (no NEW findings), preview, render ----
        if ctx.work and os.path.exists(ctx.work):
            validation = ctx.validate()
            report["validation"] = {"passed": validation["passed"], "errors": validation["errorCount"],
                                    "warnings": validation["warningCount"]}
            if baseline is not None and not entry["allow_new_findings"]:
                before = {finding_signature(f) for f in baseline["validation"]["findings"]}
                new = [f for f in validation["findings"] if finding_signature(f) not in before]
                ctx.expect(not new, f"no NEW validation findings vs the pristine layout (got {len(new)}: "
                                    + "; ".join(f"{f['check']}: {f['message'][:60]}" for f in new[:3]) + ")")
            after_dir = os.path.join(out_dir, "after")
            preview = ctx.tool("preview_layout", {"layoutPath": ctx.work, "outputDir": after_dir})
            report["previewStats"] = preview["data"]["stats"]
            report["previewWarnings"] = [w.get("kind") for w in preview["data"].get("warnings", [])]
            if preview["data"].get("pdfPath"):
                render = ctx.server.tool(
                    "render_preview_pages",
                    {"pdfPath": preview["data"]["pdfPath"], "maxPages": entry["render_pages"]},
                    image_dir=after_dir)
                report["pages"] = f"{render['data']['pagesRendered']}/{render['data']['pageCount']}"
        report["passed"] = all(c["passed"] for c in ctx.checks) if ctx.checks else True
    except StepFailure as failure:
        report["failure"] = str(failure)
    except Exception:
        report["failure"] = traceback.format_exc(limit=4)

    report["steps"] = ctx.steps
    report["checks"] = ctx.checks
    with open(os.path.join(out_dir, "report.json"), "w") as f:
        json.dump(report, f, indent=2)
    with open(os.path.join(out_dir, "done.json"), "w") as f:
        json.dump({"hash": scenario_hash(entry), "passed": report["passed"],
                   "when": time.strftime("%Y-%m-%d %H:%M:%S")}, f, indent=2)
    return report


# ================================================================ SCENARIOS
# Table/column indexes are discovered at runtime wherever practical; where a scenario pins a
# concrete index it is corpus-verified (and the assertion failure message will say so if the
# corpus file ever changes shape).

@scenario("s01-quote-remove-discount",
          layout="StandardSalesQuote.docx",
          desc="Remove the Line Discount % column from the lines; the table must stay full-width.")
def s01(ctx: Ctx):
    table = ctx.info()["tables"][2]
    total_before = ctx.grid_total(table)
    ctx.expect(table["columnCount"] == 8, "quote lines table has the corpus-verified 8 columns")
    result = ctx.edit("remove_column", tableIndex=2, column=5)["data"]
    ctx.expect("redistributed" in result["summary"], "summary reports the width redistribution")
    after = ctx.info()["tables"][2]
    ctx.expect(after["columnCount"] == 7, "7 columns remain")
    ctx.expect(ctx.grid_total(after) == total_before, "table width unchanged (redistributed)")


@scenario("s02-po-add-line-discount",
          layout="StandardPurchaseOrder.docx",
          desc="Add a Line Discount % column (bound header label) after the last lines column, then re-fit widths.")
def s02(ctx: Ctx):
    # This scenario used to open by pinning interior insertion as a documented v1 REJECTION. That stopped
    # being true when interior-position insert_column was implemented (see CHANGELOG), but the cached pass
    # kept the stale expectation alive: scenario_hash only invalidates on a change to this function's own
    # source, so a change in the SERVER that flips a scenario's outcome went unnoticed until the cache was
    # keyed on the build too. Interior insertion now has its own coverage in TableStructureEditorTests
    # (synthetic + corpus, including a gridAfter table), so this scenario keeps to its stated purpose.
    total_before = ctx.grid_total(ctx.info()["tables"][3])
    ctx.edit("insert_column", tableIndex=3, mode="field",
             dataPath="/Purchase_Header/Purchase_Line/LineDisc_PurchLine",
             headerLabelPath="/Purchase_Header/Purchase_Line/PurchLineLineDisc_Lbl", width=1000)
    # Re-fit: take the added width back out of the (widest) description column.
    widths = ctx.info()["tables"][3]["gridColumnWidths"]
    widths[1] -= 1000
    ctx.edit("set_column_widths", tableIndex=3, widths=",".join(map(str, widths)))
    after = ctx.info()["tables"][3]
    ctx.expect(after["columnCount"] == 8, "8 columns after the append")
    ctx.expect(ctx.grid_total(after) == total_before, "table re-fitted to its original width")


@scenario("s03-subinv-header-field-and-cleanup",
          layout="SalesInvoiceForSubscriptionBilling.docx",
          desc="Add ExternalDocumentNo + caption below Payment Reference; remove the unbound "
               "shipping-agent/package-tracking controls from the same block.")
def s03(ctx: Ctx):
    ctx.edit("insert_label", label="/Header/ExternalDocumentNo_Lbl", locationType="tableCell",
             tableIndex=1, row=4, col=3)
    ctx.edit("insert_field", field="/Header/ExternalDocumentNo", locationType="tableCell",
             tableIndex=1, row=5, col=3)
    table = ctx.info()["tables"][1]
    removed = 0
    for row in table["rows"]:
        if row["rowIndex"] in (4, 5):
            for cell in row["cells"]:
                if cell["colIndex"] in (0, 1) and cell.get("controlId") is not None:
                    ctx.edit("remove_control", controlId=cell["controlId"])
                    removed += 1
    ctx.expect(removed == 4, f"all four unbound controls found and removed (got {removed})")


@scenario("s04-quote-remove-two-columns-rename-header",
          layout="StandardSalesQuote.docx",
          desc="Remove BOTH the VAT % and Line Discount % columns (descending index order), then "
               "caption the blank UOM header cell via set_cell_text.")
def s04(ctx: Ctx):
    total_before = ctx.grid_total(ctx.info()["tables"][2])
    ctx.edit("remove_column", tableIndex=2, column=6)  # VAT Pct
    ctx.edit("remove_column", tableIndex=2, column=5)  # Line Discount (index unchanged: it was left of 6)
    ctx.edit("set_cell_text", tableIndex=2, row=0, col=3, text="Unit")
    after = ctx.info()["tables"][2]
    ctx.expect(after["columnCount"] == 6, "6 columns remain after two removals")
    ctx.expect(ctx.grid_total(after) == total_before, "table width preserved across BOTH removals")
    header = next(r for r in after["rows"] if r["rowIndex"] == 0)
    ctx.expect(any(c.get("text") == "Unit" for c in header["cells"]), "UOM header cell now says 'Unit'")


@scenario("s05-po-add-two-columns-refit",
          layout="StandardPurchaseOrder.docx",
          desc="Add two columns to the PO lines (Line Discount with bound label; Item Reference with "
               "static header), then one set_column_widths re-fit for both.")
def s05(ctx: Ctx):
    total_before = ctx.grid_total(ctx.info()["tables"][3])
    ctx.edit("insert_column", tableIndex=3, mode="field",
             dataPath="/Purchase_Header/Purchase_Line/LineDisc_PurchLine",
             headerLabelPath="/Purchase_Header/Purchase_Line/PurchLineLineDisc_Lbl", width=900)
    ctx.edit("insert_column", tableIndex=3, mode="field",
             dataPath="/Purchase_Header/Purchase_Line/ItemReferenceNo_PurchLine",
             headerText="Item Reference", width=1100)
    widths = ctx.info()["tables"][3]["gridColumnWidths"]
    widths[1] -= 2000  # take both added widths back out of the description column
    ctx.expect(widths[1] > 500, "description column still has usable width after the refit")
    ctx.edit("set_column_widths", tableIndex=3, widths=",".join(map(str, widths)))
    after = ctx.info()["tables"][3]
    ctx.expect(after["columnCount"] == 9, "9 columns after both appends")
    ctx.expect(ctx.grid_total(after) == total_before, "table re-fitted to its original width")
    # The totals/summary rows below the lines are right-anchored (leading empty cells, amounts at the
    # table's right edge); the appended columns' filler cells must go to those rows' LEFT edge so the
    # summary block stays at the right edge with no border bleed under the new columns.
    totals_rows = [r for r in after["rows"]
                   if not r["isControlRow"]
                   and any("Totals" in (c.get("alias") or "") for c in r["cells"])]
    ctx.expect(bool(totals_rows), "found the right-anchored totals rows")
    ctx.expect(all(r["cells"][-1].get("alias") for r in totals_rows),
               "totals rows stay right-anchored (last cell is still the bound amount, no trailing filler)")


@scenario("s06-po-remove-then-add",
          layout="StandardPurchaseOrder.docx",
          desc="Replace a lines column: remove UOM, then append Vendor Item No - remove and insert on "
               "the same table in one session.")
def s06(ctx: Ctx):
    total_before = ctx.grid_total(ctx.info()["tables"][3])
    ctx.edit("remove_column", tableIndex=3, column=3)  # UOM
    ctx.edit("insert_column", tableIndex=3, mode="field",
             dataPath="/Purchase_Header/Purchase_Line/VendorItemNo_PurchLine",
             headerText="Vendor Item No", width=1200)
    widths = ctx.info()["tables"][3]["gridColumnWidths"]
    widths[1] -= 1200
    ctx.edit("set_column_widths", tableIndex=3, widths=",".join(map(str, widths)))
    info = ctx.info()
    after = info["tables"][3]
    ctx.expect(after["columnCount"] == 7, "7 columns: one removed, one added")
    ctx.expect(ctx.grid_total(after) == total_before, "table back at its original width")
    # The appended data cell holds an INLINE control (not a cell-level one), so it shows up in the
    # controls inventory - the DTO carries the dataset path in 'alias' ("#Nav: /...").
    ctx.expect(any("VendorItemNo_PurchLine" in (c.get("alias") or "") for c in info["controls"]),
               "the new VendorItemNo binding exists in the controls inventory")


@scenario("s07-quote-merge-split-roundtrip",
          layout="StandardSalesQuote.docx",
          desc="merge_cells three spacer-row cells into one, then split_cells back - a structural "
               "round-trip that must land exactly where it started.")
def s07(ctx: Ctx):
    def spacer_cells():
        row = next(r for r in ctx.info()["tables"][2]["rows"] if r["rowIndex"] == 1)
        return len(row["cells"])

    cells_before = spacer_cells()
    ctx.expect(cells_before == 8, "spacer row starts with the corpus-verified 8 unit cells")
    ctx.edit("merge_cells", tableIndex=2, row=1, fromColumn=0, toColumn=2)
    ctx.expect(spacer_cells() == cells_before - 2, "merge collapsed 3 cells into 1")
    ctx.edit("split_cells", tableIndex=2, row=1, cellIndex=0)
    ctx.expect(spacer_cells() == cells_before, "split restored the original cell count")


@scenario("s08-quote-label-then-field-afterControl",
          layout="StandardSalesQuote.docx",
          desc="Append a label at documentEnd, then chain an insert_field afterControl using the id "
               "the label insert returned - the id round-trip an agent relies on.")
def s08(ctx: Ctx):
    header = ctx.find_data_item(ctx.fields_root(), "Header")
    field_path, label_path = ctx.unbound_field_with_label(header)
    ctx.note(f"discovered unbound pair: {field_path} / {label_path}")
    label_result = ctx.edit("insert_label", label=label_path, locationType="documentEnd")["data"]
    ctx.edit("insert_field", field=field_path, locationType="afterControl",
             controlId=label_result["controlId"])
    controls = ctx.info()["controls"]
    ctx.expect(any((c.get("alias") or "").endswith(field_path) for c in controls),
               "the new field control is present in the layout (alias carries the dataset path)")


@scenario("s09-vatspec-remove-column-baseline-diff",
          layout="StandardSalesInvoiceVatSpec.docx",
          desc="Remove a column (with bound cells) from a NON-repeater info table in a layout that "
               "carries pre-existing corpus defects - the edit must not add a single NEW validation "
               "finding on top of them.")
def s09(ctx: Ctx):
    # Target the 3-column document-info table (Document Date | Due Date | Payment Terms):
    # corpus-verified as a plain, non-repeater table whose BOTH rows are fully bound (a label row
    # over a field row) and which uses no gridSpan/gridAfter. Removing the last column drops two
    # bound cells - remove_column's documented drop-the-binding behavior - and must add no NEW
    # findings beyond this layout's pre-existing storeItemID errors and attachedTemplate warning.
    # Deliberately NOT the widest table: this layout's lines table uses w:gridAfter, which every
    # table-structure tool still refuses (see GitHub issue #9's history).
    info = ctx.info()
    candidates = [(i, t) for i, t in enumerate(info["tables"])
                  if t["part"] == "document.xml" and t["columnCount"] == 3
                  and not any(r["isControlRow"] for r in t["rows"])]
    ctx.expect(len(candidates) == 1, "exactly one 3-column non-repeater body table found")
    index, table = candidates[0]
    bound = sum(1 for r in table["rows"] for c in r["cells"] if c.get("controlId") is not None)
    ctx.expect(bound == 6, "all six cells of the info table are bound (labels over fields)")
    total_before = ctx.grid_total(table)
    ctx.edit("remove_column", tableIndex=index, column=table["columnCount"] - 1)
    after = ctx.info()["tables"][index]
    ctx.expect(after["columnCount"] == table["columnCount"] - 1, "one column removed")
    ctx.expect(ctx.grid_total(after) == total_before, "table width preserved")


@scenario("s10-statement-remove-aging-column",
          layout="StandardStatement.docx",
          desc="Remove the last aging-band column from the statement's widest lines table and keep "
               "the width; the 10-page heavyweight layout.")
def s10(ctx: Ctx):
    info = ctx.info()
    candidates = [(i, t) for i, t in enumerate(info["tables"])
                  if t["part"] == "document.xml" and any(r["isControlRow"] for r in t["rows"])]
    index, table = max(candidates, key=lambda pair: pair[1]["columnCount"])
    ctx.note(f"widest repeater table: index {index} with {table['columnCount']} columns")
    total_before = ctx.grid_total(table)
    ctx.edit("remove_column", tableIndex=index, column=table["columnCount"] - 1)
    after = ctx.info()["tables"][index]
    ctx.expect(after["columnCount"] == table["columnCount"] - 1, "one aging-band column removed")
    ctx.expect(ctx.grid_total(after) == total_before, "table width preserved")


@scenario("s11-quote-insert-repeater-table",
          layout="StandardSalesQuote.docx",
          desc="Author a NEW repeater table at documentEnd bound to /Header/Line with four columns "
               "and explicit widths.")
def s11(ctx: Ctx):
    tables_before = len(ctx.info()["tables"])
    ctx.edit("insert_repeater_table", dataItem="/Header/Line",
             columns="ItemNo_Line,Description_Line,Quantity_Line,LineAmount_Line",
             locationType="documentEnd", columnWidths="1500,4500,1500,2706")
    info = ctx.info()
    ctx.expect(len(info["tables"]) == tables_before + 1, "one new table appeared")
    new_table = info["tables"][tables_before - 4]  # body tables come before header/footer parts
    ctx.note(f"new table columns: {new_table['columnCount']}")


@scenario("s12-create-layout-and-author",
          layout=None,
          desc="Full authoring flow: create_layout from the quote's schema, then insert a repeater "
               "table plus a field+label pair into the fresh document.")
def s12(ctx: Ctx):
    ctx.work = os.path.join(ctx.out_dir, "work.docx")
    created = ctx.tool("create_layout", {
        "schemaSource": os.path.join(CORPUS_DIR, "StandardSalesQuote.docx"),
        "outputPath": ctx.work})["data"]
    ctx.expect(created["quickValidation"]["passed"], "created layout passes quick validation")
    ctx.edit("insert_label", label="/Header/YourReference__Lbl", locationType="documentEnd")
    ctx.edit("insert_field", field="/Header/YourReference", locationType="documentEnd")
    ctx.edit("insert_repeater_table", dataItem="/Header/Line",
             columns="ItemNo_Line,Description_Line,Quantity_Line,UnitPrice,LineAmount_Line",
             locationType="documentEnd")
    validation = ctx.validate()
    ctx.expect(validation["errorCount"] == 0, "authored-from-scratch layout validates with zero errors")


@scenario("s13-quote-refresh-xml-part",
          layout="StandardSalesQuote.docx",
          desc="refresh_xml_part twice: a same-schema no-op refresh (nothing orphaned), then a "
               "cross-report refresh against the PO schema (orphans reported, bindings left alone).",
          allow_new_findings=True)
def s13(ctx: Ctx):
    noop = ctx.edit("refresh_xml_part",
                    newSchemaSource=os.path.join(CORPUS_DIR, "StandardSalesQuote.docx"))["data"]
    ctx.expect(not noop.get("namespaceChanged"), "same-schema refresh: namespace unchanged")
    # The quote carries ONE pre-existing corpus defect: a binding to CompanyABNNumber, which is
    # absent from its own schema - a same-schema refresh correctly reports exactly that one binding
    # as orphaned (it cannot resolve against the new schema either), and nothing else.
    orphans = json.dumps(noop.get("orphanedBindings") or [])
    ctx.expect("CompanyABNNumber" in orphans if (noop.get("orphanedBindings") or []) else True,
               "same-schema refresh: only the pre-existing CompanyABNNumber defect is orphaned")
    ctx.expect(len(noop.get("orphanedBindings") or []) <= 1,
               "same-schema refresh orphans nothing beyond the known corpus defect")
    ctx.expect(not noop.get("newUnboundFields"), "same-schema refresh: empty old-vs-new diff")

    crossed = ctx.edit("refresh_xml_part",
                       newSchemaSource=os.path.join(CORPUS_DIR, "StandardPurchaseOrder.docx"))["data"]
    ctx.expect(crossed.get("namespaceChanged") is True, "cross-report refresh flips the namespace")
    ctx.expect(len(crossed.get("orphanedBindings") or []) > 10,
               "cross-report refresh reports the quote's bindings as orphans")
    ctx.note(f"orphans: {len(crossed.get('orphanedBindings') or [])}, "
             f"remapped: {crossed.get('remappedCount')}")


@scenario("s14-po-footer-field",
          layout="StandardPurchaseOrder.docx",
          desc="Insert a bound field into a FOOTER part (layoutPart='footer') and confirm it lands in "
               "the footer, not the body.")
def s14(ctx: Ctx):
    header_item = ctx.find_data_item(ctx.fields_root(), "Purchase_Header")
    field_path, _ = ctx.unbound_field_with_label(header_item)
    ctx.note(f"discovered unbound header field: {field_path}")
    result = ctx.edit("insert_field", field=field_path, locationType="documentEnd",
                      layoutPart="footer")["data"]
    ctx.expect(result["part"].startswith("footer"), f"control landed in a footer part ({result['part']})")
    controls = ctx.info()["controls"]
    ctx.expect(any((c.get("alias") or "").endswith(field_path) and (c.get("part") or "").startswith("footer")
                   for c in controls),
               "get_layout_info sees the new control in the footer part")


@scenario("s15-quote-rebind-header-cell",
          layout="StandardSalesQuote.docx",
          desc="Replace a bound header field in place: remove_control keepText=false empties the cell "
               "(column preserved), then insert_field binds a different column into the same cell; a "
               "second control is removed keepText=true (text kept, binding dropped).")
def s15(ctx: Ctx):
    # Corpus quirk this scenario deliberately keeps: cell 0's field (YourReference) is an INLINE
    # control inside the cell (innerControlIds), while cell 1's (QuoteValidToDate) is CELL-LEVEL
    # (controlId) - so the two removals also cover both control levels.
    table = ctx.info()["tables"][1]
    field_row = next(r for r in table["rows"] if r["rowIndex"] == 1)
    first, second = field_row["cells"][0], field_row["cells"][1]
    first_id = first.get("controlId") or (first.get("innerControlIds") or [None])[0]
    second_id = second.get("controlId") or (second.get("innerControlIds") or [None])[0]
    ctx.expect(first_id is not None and second_id is not None,
               "corpus-verified: header info block row 1 cells 0/1 carry bound controls")

    ctx.edit("remove_control", controlId=first_id)                  # removes control and content
    ctx.edit("remove_control", controlId=second_id, keepText=True)  # keeps the text

    header_item = ctx.find_data_item(ctx.fields_root(), "Header")
    field_path, _ = ctx.unbound_field_with_label(header_item)
    ctx.note(f"rebinding emptied cell to: {field_path}")
    ctx.edit("insert_field", field=field_path, locationType="tableCell", tableIndex=1, row=1, col=0)

    after_row = next(r for r in ctx.info()["tables"][1]["rows"] if r["rowIndex"] == 1)
    ctx.expect(len(after_row["cells"]) == len(field_row["cells"]),
               "cell/column count untouched by the control swaps")
    cell0, cell1 = after_row["cells"][0], after_row["cells"][1]
    ctx.expect(cell0.get("controlId") is not None or bool(cell0.get("innerControlIds")),
               "cell 0 is bound again")
    ctx.expect(cell1.get("controlId") is None and not cell1.get("innerControlIds"),
               "cell 1 is plain text now (binding dropped, text kept)")


@scenario("s16-inventory-remove-column-labels-classified",
          layout="InventoryOrderDetails.docx",
          desc="Remove a column from the <Labels>-convention layout; width preserved, the <Labels> "
               "columns classify as labels out of the box, and no labels-convention-hint fires.")
def s16(ctx: Ctx):
    info = ctx.info()
    candidates = [(i, t) for i, t in enumerate(info["tables"])
                  if t["part"] == "document.xml" and any(r["isControlRow"] for r in t["rows"])]
    index, table = max(candidates, key=lambda pair: pair[1]["columnCount"])
    total_before = ctx.grid_total(table)
    ctx.edit("remove_column", tableIndex=index, column=table["columnCount"] - 1)
    after = ctx.info()["tables"][index]
    ctx.expect(ctx.grid_total(after) == total_before, "table width preserved")
    # The default convention's labels-data-item rule classifies this layout's <Labels> columns with no
    # configuration, so previews sample them as captions and the hint has nothing to point at (it only
    # fires when a host explicitly disables/retargets the rule via BCWL_LABELS_DATA_ITEM).
    labels_item = ctx.find_data_item(ctx.fields_root(), "Labels")
    ctx.expect(labels_item is not None and bool(labels_item["columns"])
               and all(c["isLabel"] for c in labels_item["columns"]),
               "every direct column of the <Labels> data item classifies as a label by default")
    preview = ctx.tool("preview_layout", {"layoutPath": ctx.work,
                                          "outputDir": os.path.join(ctx.out_dir, "hint-check")})
    kinds = [w.get("kind") for w in preview["data"].get("warnings", [])]
    ctx.expect("labels-convention-hint" not in kinds,
               f"no labels-convention-hint under the default convention (got {kinds})")


@scenario("s17-quote-page-number-in-footer",
          layout="StandardSalesQuote.docx",
          desc="Compose the stock 'Page X / Y' idiom in the default footer: Page_Lbl label, two literal "
               "spaces, then insert_page_number's PAGE/NUMPAGES field codes (issue #29).")
def s17(ctx: Ctx):
    label = ctx.edit("insert_label", label="/Header/Page_Lbl", locationType="documentEnd",
                     layoutPart="footer")["data"]
    ctx.expect(label["part"].startswith("footer"),
               f"label resolved into a footer part ({label['part']})")

    # Each afterControl insert lands IMMEDIATELY after its anchor, so the inline sequence is built
    # outside-in: the fields first, then the separator spaces between label and fields — the same
    # anchoring rule the skill documents for any label/separator/value line.
    fields = ctx.edit("insert_page_number", locationType="afterControl", controlId=label["controlId"],
                      layoutPart="footer")["data"]
    ctx.expect(fields["controlId"] == 0, "field codes are plain runs - no controlId to address")
    ctx.expect(fields["part"] == label["part"], "fields landed in the same footer part as the label")
    ctx.expect("PAGE / NUMPAGES" in fields["summary"], "summary names the emitted construct")
    ctx.edit("insert_text", text="  ", locationType="afterControl", controlId=label["controlId"],
             layoutPart="footer")


@scenario("s18-quote-totals-rows-inside-lines-table",
          layout="StandardSalesQuote.docx",
          desc="Append the stock inside-the-table totals shape to the quote lines table: a spacer row, "
               "then a bold ruled grand-total row bound to two unused Totals columns (issue #28).")
def s18(ctx: Ctx):
    before = ctx.info()["tables"][2]
    ctx.expect(before["columnCount"] == 8, "quote lines table has the corpus-verified 8 columns")
    rows_before = before["rowCount"]

    ctx.edit("insert_table_row", tableIndex=2, cells="8:-")  # the stock spacer row
    result = ctx.edit("insert_table_row", tableIndex=2,
                      cells="-,-,-,-,3:/Header/Totals/TotalExcludingVATText,/Header/Totals/TotalSubTotal",
                      alignments="-,-,-,-,-,right", bold=True)["data"]
    ctx.expect("renders exactly once" in result["summary"],
               "summary pins the static (non-repeating) semantics")

    after = ctx.info()["tables"][2]
    ctx.expect(after["rowCount"] == rows_before + 2, "both rows joined the lines table itself")
    last = next(r for r in after["rows"] if r["rowIndex"] == after["rowCount"] - 1)
    ctx.expect(any("TotalSubTotal" in (c.get("text") or "") for c in last["cells"]),
               "the grand-total amount cell is bound in the table's last row")

    # The rule above the totals block - the remaining half of the stock look.
    ctx.edit("set_cell_borders", tableIndex=2, row=after["rowCount"] - 1, edges="top")


@scenario("s19-subinv-grouped-lines-with-subtotals",
          layout="SalesInvoiceForSubscriptionBilling.docx",
          desc="Build the grouped list shape from scratch: /Header/Line repeater, nested AssemblyLine "
               "detail row, per-group spacer + bold subtotal rows (issue #30), closed by a static "
               "grand-total row (issue #28).")
def s19(ctx: Ctx):
    table = ctx.edit("insert_repeater_table", dataItem="/Header/Line",
                     columns="ItemNo_Line,Description_Line,Quantity_Line,LineAmount_Line",
                     locationType="documentEnd", columnAlignments="left,left,right,right")["data"]

    ctx.edit("insert_repeater_row", parentControlId=table["controlId"],
             dataItem="/Header/Line/AssemblyLine", cells="-,2:Description_AssemblyLine,-")

    # The corpus (SalespersonCommission) group order: spacer row first, bold subtotal row second.
    ctx.edit("insert_subtotal_row", parentControlId=table["controlId"], cells="4:-")
    subtotal = ctx.edit("insert_subtotal_row", parentControlId=table["controlId"],
                        cells="2:-,/Header/Line/AmountExcludingVAT_Line_Lbl,/Header/Line/AmountExcludingVAT_Line",
                        alignments="-,right,right", bold=True)["data"]
    ctx.expect(subtotal["controlId"] == 0, "a static row carries no controlId of its own")
    ctx.expect("once per group" in subtotal["summary"], "summary pins the per-group semantics")

    grand = ctx.edit("insert_table_row", tableIndex=table["tableIndex"],
                     cells="2:-,/Header/Totals/TotalIncludingVATText,/Header/Totals/TotalAmountIncludingVAT",
                     alignments="-,right,right", bold=True)["data"]
    ctx.expect("renders exactly once" in grand["summary"], "the grand total is table-static, not per-group")

    after = ctx.info()["tables"][table["tableIndex"]]
    ctx.edit("set_cell_borders", tableIndex=table["tableIndex"], row=after["rowCount"] - 1, edges="top")


# ================================================================ CLI

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("names", nargs="*", help="scenario name prefixes to run (default: all)")
    parser.add_argument("--force", action="store_true", help="rerun even if cached as passed")
    parser.add_argument("--list", action="store_true", help="list scenarios and cached status")
    args = parser.parse_args()

    selected = [s for s in SCENARIOS
                if not args.names or any(s["name"].startswith(p) for p in args.names)]
    if args.names and not selected:
        print(f"no scenario matches {args.names}", file=sys.stderr)
        return 2

    if args.list:
        for s in SCENARIOS:
            done_path = os.path.join(OUT_ROOT, s["name"], "done.json")
            status = "not run"
            if os.path.exists(done_path):
                with open(done_path) as f:
                    done = json.load(f)
                fresh = done.get("hash") == scenario_hash(s)
                status = ("PASSED" if done.get("passed") else "FAILED") + \
                         ("" if fresh else " (code changed - stale)")
            print(f"{s['name']:45s} {status:28s} {s['desc'][:80]}")
        return 0

    if not os.path.exists(HOST):
        print(f"server not built: {HOST}\nrun: dotnet build -c Release", file=sys.stderr)
        return 2

    server = McpServer()
    failures = 0
    try:
        for entry in selected:
            done_path = os.path.join(OUT_ROOT, entry["name"], "done.json")
            if not args.force and os.path.exists(done_path):
                with open(done_path) as f:
                    done = json.load(f)
                if done.get("passed") and done.get("hash") == scenario_hash(entry):
                    print(f"SKIP  {entry['name']} (passed {done['when']})")
                    continue
            print(f"RUN   {entry['name']} ...", flush=True)
            report = run_scenario(server, entry)
            if report["passed"]:
                print(f"PASS  {entry['name']}  pages={report.get('pages', '-')} "
                      f"checks={len(report['checks'])}")
            else:
                failures += 1
                print(f"FAIL  {entry['name']}: {report.get('failure', 'a check failed')[:200]}")
    finally:
        server.close()

    print(f"\n{len(selected)} selected, {failures} failed -> {OUT_ROOT}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
