"""Builds the BC-sandbox fidelity pack: layouts a human uploads to a real Business Central sandbox.

This is the tooling half of backlog item B17 (and B42) - the one assumption under the whole preview
pillar that this repo cannot verify on its own: no tool-emitted layout has ever been rendered by the
real BC report engine. The pack makes that check a mechanical exercise instead of a project:

    python tools/e2e/sandbox_pack.py            # build every pack item
    python tools/e2e/sandbox_pack.py --list     # list them
    python tools/e2e/sandbox_pack.py p01 p04    # build specific items (prefix match)

Each item produces, under preview-output/sandbox-pack/<id>-<name>/ (gitignored):

    <ReportName>.docx     the layout to upload to BC
    mock-preview.pdf      this tool's OFFLINE mock render of the same layout
    mock-pages/page-N.png page images of that mock, for quick side-by-side
    README.md             report id, what was changed, and what to look at in the BC render
    report.json           every tool call, the full validation result, merge stats

Plus, at the pack root, INSTRUCTIONS.md (the BC-side procedure) and COMPARISON.md (the sheet to fill
in). The point of every item is a QUESTION about the real renderer, not a demo - see each item's
`asks` list, which flows into both README.md and COMPARISON.md.

The server is the Release build; build it first (dotnet build -c Release).
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scenarios import CORPUS_DIR, REPO, McpServer, finding_signature  # noqa: E402

OUT_ROOT = os.path.join(REPO, "preview-output", "sandbox-pack")
PREVIEW_ROWS = 4
RENDER_PAGES = 3

# Rows a repeater gets in the MOCK preview. BC decides its own row count from real data, so this only
# affects how much of the mock a reviewer sees - it is not a fidelity dimension.


# ---------------------------------------------------------------- build context

class ItemFailure(Exception):
    pass


class Build:
    """Tool sugar for one pack item, recording every call for report.json."""

    def __init__(self, server: McpServer, work: str):
        self.server = server
        self.work = work
        self.steps: list[dict] = []
        self.changes: list[str] = []

    def tool(self, name: str, args: dict) -> dict:
        env = self.server.tool(name, args)
        data = env.get("data") if isinstance(env.get("data"), dict) else None
        self.steps.append({"tool": name, "args": args, "ok": env.get("ok"),
                           "summary": (data or {}).get("summary"), "error": env.get("error")})
        if not env.get("ok"):
            raise ItemFailure(f"{name} failed: {json.dumps(env.get('error'))[:400]}")
        return env

    def edit(self, name: str, **kwargs) -> dict:
        return self.tool(name, {"layoutPath": self.work, **kwargs})

    def info(self) -> dict:
        return self.tool("get_layout_info", {"layoutPath": self.work})["data"]

    def dataset(self) -> dict:
        return self.tool("list_dataset_fields", {"source": self.work})["data"]["root"]

    def note(self, change: str):
        self.changes.append(change)

    # -- dataset helpers ---------------------------------------------------

    @staticmethod
    def find_item(node: dict, name: str) -> dict | None:
        if node.get("name") == name:
            return node
        for child in node.get("children", []):
            hit = Build.find_item(child, name)
            if hit:
                return hit
        return None

    @staticmethod
    def paths(node: dict) -> set[str]:
        out = {c["path"] for c in node.get("columns", [])}
        for child in node.get("children", []):
            out |= Build.paths(child)
        return out

    @staticmethod
    def pick(available: set[str], *candidates: str) -> str:
        """First candidate path that actually exists in this dataset - the corpus captures differ
        between reports, and a guessed path is a failed build 20 tool calls later."""
        for c in candidates:
            if c in available:
                return c
        raise ItemFailure(f"none of {candidates} exist in the dataset")

    @staticmethod
    def repeater(info: dict, data_item_suffix: str) -> dict:
        for c in info["controls"]:
            if str(c.get("kind", "")).lower().startswith("repeat") \
                    and str(c.get("alias", "")).endswith(data_item_suffix):
                return c
        raise ItemFailure(f"no repeater for data item ending {data_item_suffix!r}")

    def clear_cell(self, table_index: int, row: int, col: int) -> int:
        """Strip every control out of one cell so it becomes a plain-text cell.

        A cell can hold its control two ways: as a CELL-level sdt (`controlId`, the shape the stock
        corpus headers use) or as an INLINE sdt inside the paragraph (`innerControlIds`, the shape
        insert_column's header cell produces). set_cell_text refuses both, so both have to go.
        """
        info = self.info()
        table = next(t for t in info["tables"]
                     if t["tableIndex"] == table_index and t["part"] == "document.xml")
        cell = next(c for c in table["rows"][row]["cells"] if c["colIndex"] == col)
        ids = ([cell["controlId"]] if cell.get("controlId") is not None else []) \
            + list(cell.get("innerControlIds") or [])
        for cid in ids:
            self.edit("remove_control", controlId=cid)
        return len(ids)

    @staticmethod
    def line_table(info: dict, repeater: dict) -> int:
        idx = repeater.get("tableIndex")
        if idx is None:
            raise ItemFailure("repeater is not in a table")
        return idx


# ---------------------------------------------------------------- pack items

ITEMS: list[dict] = []


def item(id: str, title: str, report: str, base: str | None, purpose: str,
         asks: list[str], bc_steps: list[str], env: dict | None = None,
         schema: str | None = None):
    def register(fn):
        ITEMS.append({"id": id, "title": title, "report": report, "base": base, "schema": schema,
                      "purpose": purpose, "asks": asks, "bc_steps": bc_steps, "env": env or {},
                      "fn": fn})
        return fn
    return register


@item(
    id="p01", title="Quote authored entirely by the tools", report="1304 - Standard Sales Quote",
    base=None, schema="StandardSalesQuote.docx",
    purpose="A layout built from nothing but create_layout + insert_* calls - no hand-authored OOXML, "
            "no base-app layout underneath. This is the single biggest unverified claim in the README.",
    asks=[
        "Does BC accept and render a layout this tool created from scratch at all?",
        "Do the header fields, the logo placeholder and the line repeater all carry real data?",
        "Is the line table's column alignment/width sane in the real renderer?",
    ],
    bc_steps=[
        "Sales quote: pick any open quote in Cronus (Sales > Quotes) that has at least 3 lines.",
        "Report Layouts > New layout for report 1304, upload this .docx, then run the quote's Print/Send > Print with this layout selected.",
    ],
)
def p01(b: Build):
    ds = b.dataset()
    header = Build.find_item(ds, "Header") or ItemFailure("no Header data item")
    all_paths = Build.paths(ds)

    # Logo, then a two-column address block: company on the left, customer on the right.
    b.edit("insert_picture", picture=Build.pick(all_paths, "/Header/CompanyPicture"),
           locationType="documentEnd", widthMm=35, heightMm=18)
    b.note("logo PICTURE control bound to /Header/CompanyPicture")

    addr = b.edit("insert_table", rows=6, columns=2, locationType="documentEnd",
                  columnWidths="4600,4600")["data"]
    addr_idx = addr["tableIndex"]
    for row in range(6):
        b.edit("insert_field", field=f"/Header/CompanyAddress{row + 1}", locationType="tableCell",
               tableIndex=addr_idx, row=row, col=0)
        b.edit("insert_field", field=f"/Header/CustomerAddress{row + 1}", locationType="tableCell",
               tableIndex=addr_idx, row=row, col=1)
    b.note("2x6 unbound table filled with CompanyAddress1-6 / CustomerAddress1-6 FIELD controls")

    # Document info: label/value pairs in a small borderless grid.
    info_tbl = b.edit("insert_table", rows=3, columns=2, locationType="documentEnd",
                      columnWidths="2800,3000")["data"]
    info_idx = info_tbl["tableIndex"]
    pairs = [
        (Build.pick(all_paths, "/Header/DocumentDate_Lbl"), Build.pick(all_paths, "/Header/DocumentDate")),
        (Build.pick(all_paths, "/Header/QuoteValidToDate_Lbl"), Build.pick(all_paths, "/Header/QuoteValidToDate")),
        (Build.pick(all_paths, "/Header/SelltoCustomerNo_Lbl", "/Header/BilltoCustomerNo_Lbl"),
         Build.pick(all_paths, "/Header/SelltoCustomerNo", "/Header/BilltoCustumerNo")),
    ]
    for row, (label, field) in enumerate(pairs):
        b.edit("insert_label", label=label, locationType="tableCell",
               tableIndex=info_idx, row=row, col=0)
        b.edit("insert_field", field=field, locationType="tableCell",
               tableIndex=info_idx, row=row, col=1)
    b.note("3x2 label/value grid: document date, valid-to date, customer no (LABEL + FIELD controls)")

    # The line items repeater - the flagship tool.
    rep = b.edit("insert_repeater_table", dataItem="/Header/Line",
                 columns="ItemNo_Line,Description_Line,Quantity_Line,UnitPrice,LineAmount_Line",
                 locationType="documentEnd", headerFromLabels=True,
                 columnWidths="1600,4000,1100,1400,1600",
                 columnAlignments="left,left,right,right,right")["data"]
    b.note(f"repeater table over /Header/Line, {rep['columnCount']} columns, header bound to the *_Lbl label columns")

    # Totals: a right-anchored two-column block with a rule above it, the BC-native shape.
    totals = b.edit("insert_table", rows=2, columns=2, locationType="documentEnd",
                    columnWidths="6200,2400", columnAlignments="right,right")["data"]
    tot_idx = totals["tableIndex"]
    b.edit("insert_field", field=Build.pick(all_paths, "/Header/Totals/TotalIncludingVATText"),
           locationType="tableCell", tableIndex=tot_idx, row=0, col=0)
    b.edit("insert_field", field=Build.pick(all_paths, "/Header/Totals/TotalAmountIncludingVAT"),
           locationType="tableCell", tableIndex=tot_idx, row=0, col=1, bold=True)
    b.edit("set_cell_borders", tableIndex=tot_idx, row=0, edges="top", size=4)
    b.note("right-anchored totals block with a 1/2pt rule above it (set_cell_borders)")


@item(
    id="p02", title="Line table restructured on a real base-app layout",
    report="1306 - Standard Sales Invoice", base="SalesInvoiceForSubscriptionBilling.docx",
    purpose="The table-structure tools (insert_column, set_column_widths, set_cell_text, "
            "set_cell_borders) applied to the line-items table of a stock layout.",
    asks=[
        "Does the added Unit of Measure column carry data in BC, in the right place?",
        "Are the re-widened columns proportioned in the BC render the way the mock shows them?",
        "Does the rule under the header row survive as a rule (not a full grid)?",
        "The added column's header is a STATIC text cell while every other header is a bound label - "
        "do they render identically (font, weight, alignment)?",
        "insert_column binds a header with an INLINE (paragraph-level) control, where the stock header "
        "cells use CELL-level controls. Both are valid OOXML; does BC treat them the same? Compare the "
        "added column's data cells against the stock ones.",
        "insert_column adds a bound cell to EVERY row, including rows outside the repeating section - "
        "the mock reports 3 'xpath-fallback' warnings for exactly that. What does BC put in those "
        "non-repeating rows: the first line's value, blank, or an error?",
    ],
    bc_steps=[
        "Posted sales invoice: Sales > Posted Sales Invoices, pick one with several lines.",
        "Report Layouts > New layout for report 1306, upload, then Print the posted invoice with it.",
    ],
)
def p02(b: Build):
    info = b.info()
    rep = Build.repeater(info, "/Header/Line")
    idx = Build.line_table(info, rep)
    table = next(t for t in info["tables"] if t["tableIndex"] == idx and t["part"] == "document.xml")
    before = table["columnCount"]

    b.edit("insert_column", tableIndex=idx, mode="field", dataPath="/Header/Line/UnitOfMeasure",
           headerLabelPath="/Header/Line/UnitOfMeasure_Lbl", atColumn=3, width=900)
    b.note(f"insert_column: bound Unit of Measure column added at grid position 3 (table had {before} columns)")

    info = b.info()
    table = next(t for t in info["tables"] if t["tableIndex"] == idx and t["part"] == "document.xml")
    widths = list(table["gridColumnWidths"])
    total = sum(widths)
    # Widen the description column at the expense of the numeric ones, keeping the grid total intact
    # (a table whose columns no longer sum to the original width reflows the whole page in Word).
    widths[1] = int(widths[1] * 1.25)
    drift = sum(widths) - total
    widths[-1] -= drift
    b.edit("set_column_widths", tableIndex=idx, widths=",".join(str(w) for w in widths))
    b.note(f"set_column_widths: description column widened 25%, taken off the last column, grid total unchanged ({total} twips)")

    # Re-label the new column's header from a bound label to static text. Every header cell here is a
    # control cell of one shape or another, and set_cell_text refuses those by design - so the
    # BC-authoring pattern is remove the control first, then set the text.
    b.clear_cell(idx, row=0, col=3)
    b.edit("set_cell_text", tableIndex=idx, row=0, col=3, text="UoM")
    b.note("remove_control + set_cell_text: the new column's header swapped from a bound label to "
           "static 'UoM' text (so the render shows a bound header and a static one side by side)")

    b.edit("set_cell_borders", tableIndex=idx, row=0, edges="bottom", size=4)
    b.note("set_cell_borders: 1/2pt rule under the header row")


@item(
    id="p03", title="Fields bound into a header and a footer",
    report="1322 - Standard Purchase Order", base="StandardPurchaseOrder.docx",
    purpose="Bound controls placed in the header and footer PARTS rather than the body - the "
            "location machinery no other pack item exercises.",
    asks=[
        "WHICH header part does BC use on page 1? Each part now carries its own marker "
        "([H1]/[H2]/[H3], [F1]/[F2]/[F3]) - whichever appears is the part BC rendered.",
        "Round 2 showed NOTHING: the insertion went to header2.xml (the section's DEFAULT header) but "
        "this layout sets w:titlePg, so a one-page document renders header3.xml (the FIRST-page header). "
        "This round marks every part to prove that reading.",
        "On a 2+ page order, do the default-header and default-footer markers appear from page 2 on?",
        "Does each marked field still carry its data, and does the static separator text sit inline?",
    ],
    bc_steps=[
        "Purchase order: Purchasing > Purchase Orders, pick one with enough lines to run to 2+ pages "
        "(add lines if needed - the multi-page behaviour is the point, and round 1 rendered on one page "
        "so it could not answer it).",
        "Report Layouts > New layout for report 1322, upload, then Print the order with it.",
    ],
)
def p03(b: Build):
    ds = b.dataset()
    all_paths = Build.paths(ds)
    no_field = Build.pick(all_paths, "/Purchase_Header/No_PurchHeader")
    date_field = Build.pick(all_paths, "/Purchase_Header/DocumentDate")
    phone = Build.pick(all_paths, "/Purchase_Header/CompanyPhoneNo")

    # Mark EVERY header/footer part, each with its own tag. Round 2 targeted `layoutPart="header"` with
    # no partName, which resolves to the section's DEFAULT header (header2.xml) - but this layout sets
    # w:titlePg, so a one-page render uses the FIRST-page header (header3.xml) and the insertion was
    # invisible. Marking every part turns "did it land?" into "which part does BC actually render?",
    # which is the question worth answering, and it is one an agent cannot answer from get_layout_info
    # today: `parts` is a flat list of file names with no first/default/even role attached.
    parts = b.info()["parts"]
    headers = sorted(p for p in parts if p.startswith("header"))
    footers = sorted(p for p in parts if p.startswith("footer"))

    # Chaining rule, learned the hard way: `documentEnd` appends a NEW PARAGRAPH per call, so a
    # field/separator/field sequence built that way lands on three separate lines (round 3's mock proved
    # it). To keep them inline, anchor with `afterControl` - and each insert lands IMMEDIATELY after its
    # anchor, so insert both controls first (the second anchored on the first), then drop the separator
    # on the FIRST one so it slots between them.
    def marked(part: str, layout_part: str, tag: str, first: dict, second: dict, separator: str):
        id_first = b.edit(locationType="documentEnd", layoutPart=layout_part, partName=part,
                          **first)["data"]["controlId"]
        id_second = b.edit(locationType="afterControl", controlId=id_first, layoutPart=layout_part,
                           partName=part, **second)["data"]["controlId"]
        b.edit("insert_text", text=separator, locationType="afterControl", controlId=id_first,
               layoutPart=layout_part, partName=part)
        b.edit("insert_text", text=f"  {tag}", locationType="afterControl", controlId=id_second,
               layoutPart=layout_part, partName=part)

    for part in headers:
        marked(part, "header", f"[H{part.removeprefix('header').removesuffix('.xml')}]",
               {"name": "insert_field", "field": no_field},
               {"name": "insert_field", "field": date_field}, "  |  ")
    b.note(f"every header part marked and bound INLINE: {', '.join(headers)} - each ends with "
           f"'<order no>  |  <document date>  [H*]' on one line")

    phone_lbl = Build.pick(all_paths, "/Purchase_Header/CompanyPhoneNo_Lbl")
    for part in footers:
        marked(part, "footer", f"[F{part.removeprefix('footer').removesuffix('.xml')}]",
               {"name": "insert_label", "label": phone_lbl},
               {"name": "insert_field", "field": phone}, ": ")
    b.note(f"every footer part marked and bound: {', '.join(footers)}")


@item(
    id="p04", title="Nested detail row rebuilt by the tools",
    report="1016 - Job Quote", base="JobQuote.docx",
    purpose="The stock layout already nests Job Planning Lines inside Job Tasks. This item REMOVES "
            "that hand-authored nested repeater and rebuilds it with insert_repeater_row, so the BC "
            "render can be compared against the stock layout's own render of the same job - same "
            "data, our OOXML.",
    asks=[
        "Does the rebuilt detail row repeat once per planning line under each task, as the stock layout does?",
        "Do the detail cells line up with the parent grid (spans), or does the row drift?",
        "Compare against a run of the SAME job with the STOCK layout - any difference is ours.",
    ],
    bc_steps=[
        "Job/Project: Projects > Jobs, pick one with tasks that have planning lines (Cronus has several).",
        "FIRST run the report with the built-in layout and save that PDF as the reference.",
        "Then Report Layouts > New layout for report 1016, upload this .docx, and run the SAME job again.",
    ],
)
def p04(b: Build):
    info = b.info()
    inner = Build.repeater(info, "/Job/Job_Task/Job_Planning_Line")
    outer = Build.repeater(info, "/Job/Job_Task")
    b.edit("remove_control", controlId=inner["sdtId"])
    b.note("removed the layout's own hand-authored Job_Planning_Line nested repeater (and its row template)")

    # A detail row must cover the parent grid exactly, so the spec is built from the parent's real
    # column count: an indent spacer, then one cell per detail column, with the last one taking up
    # whatever slack is left.
    grid = next(t for t in info["tables"]
                if t["tableIndex"] == outer["tableIndex"] and t["part"] == outer["part"])["columnCount"]
    detail = ["Number", "Type", "Quantity", "UnitPrice", "LineDiscountPct", "TotalPrice"]
    spacer = 2
    slack = grid - spacer - len(detail)
    if slack < 0:
        detail = detail[:grid - spacer]
        slack = 0
    cells = [f"{spacer}:-"] + detail[:-1] + [f"{slack + 1}:{detail[-1]}"]
    b.edit("insert_repeater_row", parentControlId=outer["sdtId"],
           dataItem="/Job/Job_Task/Job_Planning_Line",
           cells=",".join(cells),
           alignments=",".join(["-"] + ["left", "left"] + ["right"] * (len(detail) - 2)))
    b.note(f"rebuilt it with insert_repeater_row: an indented detail row spanning the parent's "
           f"{grid}-column grid as {','.join(cells)}")


@item(
    id="p05", title="Columns removed and merged on a stock quote layout",
    report="1304 - Standard Sales Quote", base="StandardSalesQuote.docx",
    purpose="The destructive half of the table tools, which P02 does not cover: remove_column and "
            "merge_cells change the grid itself. Same report as P01, so ONE quote printed with three "
            "layouts (built-in, P01, P05) gives a from-scratch layout, an edited stock layout and the "
            "stock reference side by side.",
    asks=[
        "After a column is removed, do the remaining columns still bind to the right data in BC?",
        "Does the grid re-flow cleanly, or does the table run off the page margin?",
        "Does the merged header cell render as one cell spanning two columns?",
        "Does the rule under the header row land where the mock puts it?",
        "Round 1 of this item was REJECTED by BC ('UnknownDataBinding: .../Header/CompanyABNNumber') - "
        "a broken control the stock corpus layout already carried, which validate_layout reports as an "
        "error on the pristine file. It is now removed as the first edit. Does BC accept it this time?",
    ],
    bc_steps=[
        "Use the SAME sales quote as P01 so the two are directly comparable.",
        "Report Layouts > New layout for report 1304, upload, then Print the quote with it.",
        "Also print it with the BUILT-IN layout once - that is the reference for what was removed.",
    ],
)
def p05(b: Build):
    # The stock capture carries a CompanyABNNumber control bound to a foreign store item, which
    # validate_layout flags as 3 errors before anything is edited - and which BC rejected the layout
    # for on upload ('UnknownDataBinding'). Removing it is exactly the repair the tool's own finding
    # text prescribes, and makes this item uploadable.
    info = b.info()
    stray = [c for c in info["controls"] if "CompanyABNNumber" in str(c.get("alias", ""))]
    for c in stray:
        b.edit("remove_control", controlId=c["sdtId"])
    if stray:
        b.note(f"remove_control: dropped {len(stray)} broken CompanyABNNumber control(s) the stock "
               f"layout shipped with - the binding BC rejected the upload for")

    info = b.info()
    rep = Build.repeater(info, "/Header/Line")
    idx = Build.line_table(info, rep)
    table = next(t for t in info["tables"] if t["tableIndex"] == idx and t["part"] == "document.xml")
    before = table["columnCount"]
    total = sum(table["gridColumnWidths"])

    # Drop the VAT % column if this capture has one, else the last column - either way the grid
    # loses a column and the rest have to absorb its width.
    target = before - 1
    for cell in table["rows"][0]["cells"]:
        if "vat" in str(cell.get("alias", "")).lower():
            target = cell["colIndex"]
            break
    b.edit("remove_column", tableIndex=idx, column=target)
    b.note(f"remove_column: grid column {target} of {before} removed from the line table")

    # Give the removed column's width to the description column, so the table still spans the full
    # content width - a grid that no longer sums to the original total reflows the whole page.
    info = b.info()
    table = next(t for t in info["tables"] if t["tableIndex"] == idx and t["part"] == "document.xml")
    widths = list(table["gridColumnWidths"])
    widths[1] += total - sum(widths)
    b.edit("set_column_widths", tableIndex=idx, widths=",".join(str(w) for w in widths))
    b.note(f"set_column_widths: the freed width folded into the description column, grid total back "
           f"to {total} twips")

    # merge_cells refuses to swallow a bound control (it would silently drop the binding), so the
    # absorbed cell's label goes first - the deliberate two-step the tool's own error text prescribes.
    b.clear_cell(idx, row=0, col=1)
    b.edit("merge_cells", tableIndex=idx, row=0, fromColumn=0, toColumn=1)
    b.note("remove_control + merge_cells: the second header cell's label removed, then the first two "
           "header cells merged into one spanning cell")

    b.edit("set_cell_borders", tableIndex=idx, row=0, edges="bottom", size=4)
    b.note("set_cell_borders: 1/2pt rule under the header row")


# ---------------------------------------------------------------- runner

def build_item(entry: dict, out_dir: str) -> dict:
    os.makedirs(out_dir, exist_ok=True)
    name = entry["report"].split(" - ", 1)[-1].replace(" ", "")
    work = os.path.join(out_dir, f"{name}.docx")
    server = McpServer(env=entry["env"] or None)
    record: dict = {"id": entry["id"], "title": entry["title"], "report": entry["report"],
                    "purpose": entry["purpose"], "asks": entry["asks"], "bcSteps": entry["bc_steps"],
                    "env": entry["env"], "base": entry["base"] or f"(created from {entry['schema']})"}
    try:
        b = Build(server, work)
        if entry["base"]:
            # Validate the PRISTINE base first. Several corpus layouts ship with findings of their own
            # (StandardSalesQuote carries a stray CompanyABNNumber control bound to a foreign store
            # item, so it is 3 errors before anything is edited). Grading an item on its absolute
            # finding count would fail it for defects it inherited; only NEW findings are ours.
            base_env = b.tool("validate_layout",
                              {"layoutPath": os.path.join(CORPUS_DIR, entry["base"]),
                               "level": "full"})["data"]
            record["baseValidation"] = {"errorCount": base_env["errorCount"],
                                        "warningCount": base_env["warningCount"]}
            base_sigs = {finding_signature(f) for f in base_env["findings"]}
            shutil.copy2(os.path.join(CORPUS_DIR, entry["base"]), work)
        else:
            base_sigs = set()
            # The heading lands on the printed document, so it reads as the business document it is,
            # not as the name of a test item.
            b.tool("create_layout", {"schemaSource": os.path.join(CORPUS_DIR, entry["schema"]),
                                     "outputPath": work,
                                     "headingText": entry["report"].split(" - ", 1)[-1].upper()})
        entry["fn"](b)

        validation = b.tool("validate_layout", {"layoutPath": work, "level": "full"})["data"]
        preview = b.tool("preview_layout", {"layoutPath": work, "rows": PREVIEW_ROWS,
                                            "outputDir": out_dir})["data"]
        new_findings = [f for f in validation["findings"]
                        if finding_signature(f) not in base_sigs]
        record.update({"changes": b.changes, "steps": b.steps, "validation": validation,
                       "newFindings": new_findings,
                       "preview": {k: preview.get(k) for k in
                                   ("mergedDocxPath", "pdfPath", "converterUsed", "conversionOk",
                                    "conversionError", "stats", "warnings")}})
        if preview.get("pdfPath"):
            target = os.path.join(out_dir, "mock-preview.pdf")
            if os.path.abspath(preview["pdfPath"]) != os.path.abspath(target):
                shutil.move(preview["pdfPath"], target)   # move, not copy: one PDF per item, named
                                                          # predictably, so the folder stays obvious
            server.tool("render_preview_pages", {"pdfPath": target, "maxPages": RENDER_PAGES},
                        image_dir=os.path.join(out_dir, "mock-pages"))
        record["ok"] = not any(f["severity"] == "Error" for f in new_findings)
        record["layoutFile"] = os.path.basename(work)
    except Exception as exc:                                  # noqa: BLE001 - recorded, not raised
        record.update({"ok": False, "failure": f"{type(exc).__name__}: {exc}",
                       "trace": traceback.format_exc()[-1500:],
                       "changes": getattr(locals().get("b", None), "changes", []),
                       "steps": getattr(locals().get("b", None), "steps", [])})
    finally:
        server.close()

    with open(os.path.join(out_dir, "report.json"), "w", encoding="utf-8") as f:
        json.dump(record, f, indent=2)
    write_item_readme(record, out_dir)
    return record


def write_item_readme(r: dict, out_dir: str):
    v = r.get("validation") or {}
    lines = [
        f"# {r['id'].upper()} - {r['title']}",
        "",
        f"**Report:** {r['report']}  ",
        f"**Layout file:** `{r.get('layoutFile', '(build failed)')}`  ",
        f"**Built from:** {r['base']}",
        "",
        "## Why this layout exists",
        "",
        r["purpose"],
        "",
        "## What the tools changed",
        "",
    ]
    lines += [f"- {c}" for c in (r.get("changes") or ["(nothing - build failed)"])]
    lines += ["", "## What to look at in the BC render", ""]
    lines += [f"{i}. {q}" for i, q in enumerate(r["asks"], 1)]
    lines += ["", "## Getting it into BC", ""]
    lines += [f"{i}. {s}" for i, s in enumerate(r["bcSteps"], 1)]
    if r.get("env"):
        lines += ["", f"> Authored with `{', '.join(f'{k}={v}' for k, v in r['env'].items())}` set on the "
                      "server. That only affects how the LAYOUT was authored - it is baked into the "
                      "file and needs no BC-side setting."]
    lines += ["", "## Mock preview (this tool's offline render - NOT a BC render)", ""]
    p = r.get("preview") or {}
    if p.get("conversionOk"):
        lines += [f"- `mock-preview.pdf` - converted with **{p.get('converterUsed')}**",
                  "- `mock-pages/page-N.png` - the same pages as images"]
    else:
        lines += [f"- No PDF: {p.get('conversionError') or 'no converter available'}",
                  "- The merged `.docx` is still there to open by hand."]
    stats = p.get("stats") or {}
    if stats:
        lines += ["", f"Merge stats: {json.dumps(stats)}"]
    base = r.get("baseValidation")
    new = r.get("newFindings")
    lines += ["", f"Validation (`validate_layout level=full`): **{v.get('errorCount', '?')} errors, "
                  f"{v.get('warningCount', '?')} warnings** in total"]
    if base:
        lines += [f"- the base layout already had {base['errorCount']} errors / {base['warningCount']} "
                  f"warnings before anything was edited",
                  f"- **{len(new or [])} finding(s) introduced by these edits**"
                  + (":" if new else " - the edits are structurally clean.")]
        lines += [f"  - `{f['severity']}` {f['check']}: {f['message'][:150]}" for f in (new or [])]
    lines += ["", "See `report.json` for every finding and every tool call.", ""]
    lines += ["> The mock merges deterministic SAMPLE data offline and converts with Word/LibreOffice.",
              "> It is not a BC render, and differences in captions, fonts and pagination are expected.",
              "> The BC render is the reference; the mock is what we are grading against it.", ""]
    if r.get("failure"):
        lines += ["## BUILD FAILED", "", f"```\n{r['failure']}\n```", ""]
    with open(os.path.join(out_dir, "README.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def write_pack_docs(records: list[dict]):
    ok = [r for r in records if r.get("ok")]
    rows = "\n".join(
        f"| {r['id'].upper()} | {r['title']} | {r['report']} | "
        f"`{r['id']}-{r['title'][:28].replace(' ', '-').lower()}/{r.get('layoutFile', '?')}` | "
        f"{'ready' if r.get('ok') else 'BUILD FAILED'} |"
        for r in records)

    instructions = f"""# BC sandbox fidelity pack - what to do

{len(ok)} of {len(records)} layouts built. Each folder holds one layout to upload to a Business Central
sandbox, this tool's own offline mock render of it, and a README saying what to look at.

**Why:** nothing this tool has ever produced has been rendered by the real BC report engine. Until that
happens, "the preview catches structural and binding mistakes" is an untested claim (backlog B17), and
the question of how BC resolves a foreign-namespace binding (B42) is open. This pack answers both.

## The layouts

| Item | What it tests | Report | File | Status |
|---|---|---|---|---|
{rows}

## Procedure, per layout

1. **Read that item's `README.md` first** - it lists the specific questions its render answers. A PDF
   with no answers attached is not much use.
2. In BC: **Report Layouts** page > filter to the report id > **New** > give it a recognisable name
   (e.g. `PACK-P01`) > upload the `.docx` from that folder.
3. Run the report as described in the item README, **selecting that layout**, and save/print to PDF.
4. Name the PDF `<item>-bc.pdf` (e.g. `p01-bc.pdf`) and drop it back in the item's folder.
5. Where the item README says so, also run the report with the **built-in** layout and save that as
   `<item>-bc-stock.pdf` - the stock render is the reference for anything we changed.

## Two things worth capturing while you are in there

- **The dataset.** If the report's request page offers it (or via *Send to > XML*), save the dataset
  XML as `<item>-dataset.xml`. With that, `preview_layout`'s `dataOverridesPath` re-runs the mock on
  the *same data as the BC render*, which turns a fuzzy visual comparison into an exact one.
- **Anything BC refuses.** If BC rejects a layout on upload, that is a finding, not a setback - save
  the exact message. Most likely cause is a dataset mismatch between the sandbox's report version and
  the corpus capture these were authored from, which `refresh_xml_part` fixes; send the message and
  the stock layout exported from your sandbox and the layout can be rebuilt against it.

## What to send back

The whole `sandbox-pack/` folder, with the BC PDFs (and any dataset XML) added. Then `COMPARISON.md`
gets filled in per item, and the real findings go to `docs/FIDELITY-CHECKLIST.md` and the backlog.
"""

    comparison = ["# Comparison sheet", "",
                  "Fill one block per item, comparing `mock-preview.pdf` against `<item>-bc.pdf`.",
                  "Dimension names match `docs/FIDELITY-CHECKLIST.md`, so findings can be copied straight",
                  "back into it. **The BC render is the reference** - every difference is the mock's to",
                  "explain, not BC's.", ""]
    for r in records:
        comparison += [f"## {r['id'].upper()} - {r['title']} ({r['report']})", ""]
        comparison += ["Questions this item exists to answer:", ""]
        comparison += [f"- [ ] {q}" for q in r["asks"]]
        comparison += ["", "| Dimension | Match? | Notes |", "|---|---|---|"]
        for dim in ("Structure renders at all", "Binding fill / field values",
                    "Caption / label text", "Repeater row count + nesting",
                    "Fonts", "Pagination / page breaks", "Number/date/locale formatting",
                    "Column widths + alignment", "Borders / rules", "Picture"):
            comparison += [f"| {dim} |  |  |"]
        comparison += ["", "Verdict: ", ""]

    os.makedirs(OUT_ROOT, exist_ok=True)
    with open(os.path.join(OUT_ROOT, "INSTRUCTIONS.md"), "w", encoding="utf-8") as f:
        f.write(instructions)
    with open(os.path.join(OUT_ROOT, "COMPARISON.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(comparison))
    with open(os.path.join(OUT_ROOT, "MANIFEST.json"), "w", encoding="utf-8") as f:
        json.dump([{k: r.get(k) for k in ("id", "title", "report", "ok", "layoutFile", "failure")}
                   for r in records], f, indent=2)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("items", nargs="*", help="item ids to build (prefix match); default all")
    ap.add_argument("--list", action="store_true")
    args = ap.parse_args()

    if args.list:
        for e in ITEMS:
            print(f"{e['id']}  {e['report']:38s} {e['title']}")
        return 0

    selected = [e for e in ITEMS if not args.items or any(e["id"].startswith(a) for a in args.items)]
    if not selected:
        print("no matching items", file=sys.stderr)
        return 2

    records = []
    for e in selected:
        slug = f"{e['id']}-{e['title'][:28].replace(' ', '-').lower()}"
        out_dir = os.path.join(OUT_ROOT, slug)
        print(f"--- {e['id']}: {e['title']}")
        r = build_item(e, out_dir)
        r["slug"] = slug
        records.append(r)
        v = r.get("validation") or {}
        print(f"    {'OK ' if r.get('ok') else 'FAIL'} new={len(r.get('newFindings') or [])} errors={v.get('errorCount')} "
              f"warnings={v.get('warningCount')} pdf={'yes' if (r.get('preview') or {}).get('conversionOk') else 'no'}"
              + (f"\n    {r.get('failure')}" if r.get("failure") else ""))

    # Keep the pack-level docs describing the whole pack, not just this run's subset.
    prior = {}
    manifest = os.path.join(OUT_ROOT, "MANIFEST.json")
    if os.path.exists(manifest) and len(selected) != len(ITEMS):
        try:
            with open(manifest, encoding="utf-8") as f:
                prior = {m["id"]: m for m in json.load(f)}
        except (OSError, ValueError, KeyError):
            prior = {}
    merged = []
    for e in ITEMS:
        hit = next((r for r in records if r["id"] == e["id"]), None)
        if hit:
            merged.append(hit)
        elif e["id"] in prior:
            merged.append({**prior[e["id"]], "asks": e["asks"], "purpose": e["purpose"],
                           "bcSteps": e["bc_steps"], "title": e["title"], "report": e["report"]})
    write_pack_docs(merged)
    print(f"\npack: {OUT_ROOT}")
    return 0 if all(r.get("ok") for r in records) else 1


if __name__ == "__main__":
    sys.exit(main())
