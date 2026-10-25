"""Builds the BC-sandbox fidelity pack: layouts a human uploads to a real Business Central sandbox.

This is the tooling half of the BC-sandbox fidelity validation (and of GitHub issue #1 on
foreign-namespace bindings) - the check this repo cannot run on its own: only the real BC report
engine can prove a tool-emitted layout renders correctly. Round 1 (2026-08-01, all five items passed)
is recorded in docs/FIDELITY-CHECKLIST.md; the pack is a STANDING re-run whenever the tool surface
changes (GitHub issue #14), and it makes each round a mechanical exercise instead of a project:

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
import re
import shutil
import sys
import time
import traceback
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scenarios import CORPUS_DIR, REPO, McpServer, describe_error, finding_signature  # noqa: E402

OUT_ROOT = os.path.join(REPO, "preview-output", "sandbox-pack")
PREVIEW_ROWS = 4
RENDER_PAGES = 3

# Rows a repeater gets in the MOCK preview. BC decides its own row count from real data, so this only
# affects how much of the mock a reviewer sees - it is not a fidelity dimension.


# ---------------------------------------------------------------- table-tool arguments

def widths_and_alignments(widths, alignments=None):
    """insert_table columns: one object per column (GitHub issue #133)."""
    alignments = alignments or [None] * len(widths)
    return [{k: v for k, v in (("width", w), ("alignment", a)) if v is not None} for w, a in zip(widths, alignments)]


def repeater_columns(data_item, names, widths, alignments):
    """insert_repeater_table columns: one object per column, each field a full dataset path."""
    return [{"field": f"{data_item}/{n}", "width": w, "alignment": a} for n, w, a in zip(names, widths, alignments)]


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

    def drop_broken_abn_control(self):
        """Removes the stock quote capture's CompanyABNNumber control, bound to a foreign store item.

        validate_layout flags it as 3 errors before anything is edited, and BC rejected round 1's P05 for it
        on upload ('UnknownDataBinding'). Removing it is the repair the tool's own finding text prescribes.
        """
        stray = [c for c in self.info()["controls"] if "CompanyABNNumber" in str(c.get("alias", ""))]
        for c in stray:
            self.edit("remove_control", controlId=c["sdtId"])
        if stray:
            self.note(f"remove_control: dropped {len(stray)} broken CompanyABNNumber control(s) the stock "
                      "layout shipped with - the binding BC rejected the upload for")


# ---------------------------------------------------------------- pack items

ITEMS: list[dict] = []


def item(id: str, title: str, report: str, base: str | None, purpose: str,
         asks: list[str], bc_steps: list[str], env: dict | None = None,
         schema: str | None = None, template=None):
    """template: for a from-scratch item (base None), a function (Build, out_dir) -> path that builds the
    file create_layout gets as its templatePath."""
    def register(fn):
        ITEMS.append({"id": id, "title": title, "report": report, "base": base, "schema": schema,
                      "purpose": purpose, "asks": asks, "bc_steps": bc_steps, "env": env or {},
                      "template": template, "fn": fn})
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
        "TYPOGRAPHY (issue #3): round 1 rendered this item in a DIFFERENT font than the mock, because "
        "the layout named no font at all. Blank builds now pin Calibri 11pt in their own scaffolded "
        "styles part - do the BC render and the mock now use the SAME typeface and size?",
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

    addr = b.edit("insert_table", rowCount=6, columns=widths_and_alignments([4600, 4600]),
                  locationType="documentEnd")["data"]
    addr_idx = addr["tableIndex"]
    for row in range(6):
        b.edit("insert_field", field=f"/Header/CompanyAddress{row + 1}", locationType="tableCell",
               tableIndex=addr_idx, row=row, col=0)
        b.edit("insert_field", field=f"/Header/CustomerAddress{row + 1}", locationType="tableCell",
               tableIndex=addr_idx, row=row, col=1)
    b.note("2x6 unbound table filled with CompanyAddress1-6 / CustomerAddress1-6 FIELD controls")

    # Document info: label/value pairs in a small borderless grid.
    info_tbl = b.edit("insert_table", rowCount=3, columns=widths_and_alignments([2800, 3000]),
                      locationType="documentEnd")["data"]
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
                 columns=repeater_columns(
                     "/Header/Line", ["ItemNo_Line", "Description_Line", "Quantity_Line", "UnitPrice", "LineAmount_Line"],
                     [1600, 4000, 1100, 1400, 1600], ["left", "left", "right", "right", "right"]),
                 locationType="documentEnd", headerFromLabels=True)["data"]
    b.note(f"repeater table over /Header/Line, {rep['columnCount']} columns, header bound to the *_Lbl label columns")

    # Totals: a right-anchored two-column block with a rule above it, the BC-native shape.
    totals = b.edit("insert_table", rowCount=2, columns=widths_and_alignments([6200, 2400], ["right", "right"]),
                    locationType="documentEnd")["data"]
    tot_idx = totals["tableIndex"]
    b.edit("insert_field", field=Build.pick(all_paths, "/Header/Totals/TotalIncludingVATText"),
           locationType="tableCell", tableIndex=tot_idx, row=0, col=0)
    b.edit("insert_field", field=Build.pick(all_paths, "/Header/Totals/TotalAmountIncludingVAT"),
           locationType="tableCell", tableIndex=tot_idx, row=0, col=1, format={"bold": True})
    b.edit("set_cell_format", tableIndex=tot_idx, row=0, edges=["top"], borderSize=4)
    b.note("right-anchored totals block with a 1/2pt rule above it (set_cell_format)")


@item(
    id="p02", title="Line table restructured on a real base-app layout",
    report="1306 - Standard Sales Invoice", base="SalesInvoiceForSubscriptionBilling.docx",
    purpose="The table-structure tools (insert_column, set_column_widths, set_cell_text, "
            "set_cell_format) applied to the line-items table of a stock layout.",
    asks=[
        "Does the added Unit of Measure column carry data in BC, in the right place?",
        "Are the re-widened columns proportioned in the BC render the way the mock shows them?",
        "Does the rule under the header row survive as a rule (not a full grid)?",
        "The added column's header is a STATIC text cell while every other header is a bound label - "
        "do they render identically (font, weight, alignment)?",
        "insert_column binds a header with an INLINE (paragraph-level) control, where the stock header "
        "cells use CELL-level controls. Both are valid OOXML; does BC treat them the same? Compare the "
        "added column's data cells against the stock ones.",
        "insert_column binds the new column in the header and in the repeating line row only; every other "
        "row of the table (the totals rows) gets an empty cell or a wider trailing cell. Do those rows "
        "print with nothing in the new column, as the mock shows?",
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
    b.edit("set_column_widths", tableIndex=idx, widths=list(widths))
    b.note(f"set_column_widths: description column widened 25%, taken off the last column, grid total unchanged ({total} twips)")

    # Re-label the new column's header from a bound label to static text. Every header cell here is a
    # control cell of one shape or another, and set_cell_text refuses those by design - so the
    # BC-authoring pattern is remove the control first, then set the text.
    b.clear_cell(idx, row=0, col=3)
    b.edit("set_cell_text", tableIndex=idx, row=0, col=3, text="UoM")
    b.note("remove_control + set_cell_text: the new column's header swapped from a bound label to "
           "static 'UoM' text (so the render shows a bound header and a static one side by side)")

    b.edit("set_cell_format", tableIndex=idx, row=0, edges=["bottom"], borderSize=4)
    b.note("set_cell_format: 1/2pt rule under the header row")


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
        "Do the detail cells line up with the parent grid (spans), under the same captions as the stock row "
        "(Type, No., Quantity, Unit Price, Line Discount %, Line Discount Amount, Total Price)?",
        "Is the rebuilt row Segoe UI 8 pt like the stock rows, with no value wrapping (insert_repeater_row's "
        "format; round 3 printed it at 11 pt)?",
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
    # The stock row's own order, after its task-number and description cells.
    detail = ["Type", "Number", "Quantity", "UnitPrice", "LineDiscountPct", "LineDiscountAmount", "TotalPrice"]
    spacer = 2
    slack = grid - spacer - len(detail)
    if slack < 0:
        detail = detail[:grid - spacer]
        slack = 0
    item = "/Job/Job_Task/Job_Planning_Line"
    alignments = ["left", "left"] + ["right"] * (len(detail) - 2)
    cells = [{"span": spacer}] + [
        {"fields": [f"{item}/{name}"], "alignment": alignment}
        | ({"span": slack + 1} if i == len(detail) - 1 else {})
        for i, (name, alignment) in enumerate(zip(detail, alignments))]
    # With its only row removed, the task item has no row to take formatting from, so the stock row's
    # Segoe UI 8 pt is passed explicitly.
    b.edit("insert_repeater_row", parentControlId=outer["sdtId"], dataItem=item, cells=cells,
           format={"fontName": "Segoe UI", "fontSizePoints": 8})
    b.note(f"rebuilt it with insert_repeater_row: an indented detail row spanning the parent's "
           f"{grid}-column grid as a {spacer}-column spacer then {', '.join(detail)}, in the stock row's "
           "Segoe UI 8 pt (format)")


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
    b.drop_broken_abn_control()

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
    b.edit("set_column_widths", tableIndex=idx, widths=list(widths))
    b.note(f"set_column_widths: the freed width folded into the description column, grid total back "
           f"to {total} twips")

    # merge_cells refuses to swallow a bound control (it would silently drop the binding), so the
    # absorbed cell's label goes first - the deliberate two-step the tool's own error text prescribes.
    b.clear_cell(idx, row=0, col=1)
    b.edit("merge_cells", tableIndex=idx, row=0, fromColumn=0, toColumn=1)
    b.note("remove_control + merge_cells: the second header cell's label removed, then the first two "
           "header cells merged into one spanning cell")

    b.edit("set_cell_format", tableIndex=idx, row=0, edges=["bottom"], borderSize=4)
    b.note("set_cell_format: 1/2pt rule under the header row")


@item(
    id="p06", title="Purchase order from scratch with pinned typography", report="1322 - Standard Purchase Order",
    base=None, schema="StandardPurchaseOrder.docx",
    purpose="A second from-scratch build (issue #3's fix made blank builds ship a styles part pinning "
            "Calibri 11pt plus a TableGrid definition), exercising exactly the typography surface P01 "
            "does not: an explicit tableStyle reference and explicit format bold/size static text on "
            "top of the pinned default font.",
    asks=[
        "TYPOGRAPHY (issue #3): does the whole document render in the SAME typeface and size in BC as "
        "in the mock (Calibri 11pt from the scaffolded styles part)?",
        "The line table references the scaffolded TableGrid STYLE (w:tblStyle) rather than direct "
        "border formatting - does BC honour the style and draw the full 1/2pt grid the mock shows?",
        "The two static text runs above the lines are 16pt bold and default-size respectively "
        "(insert_text format bold/size on top of the pinned base font) - do both sizes match the mock?",
        "Do the address block, info grid, logo placeholder and totals all carry real data, as P01's did?",
    ],
    bc_steps=[
        "Purchase order: Purchasing > Purchase Orders, pick one with several lines (the same order as "
        "P03 works and makes the two directly comparable).",
        "Report Layouts > New layout for report 1322, upload this .docx, then Print the order with it.",
    ],
)
def p06(b: Build):
    ds = b.dataset()
    all_paths = Build.paths(ds)

    # Logo, then a two-column address block: company on the left, vendor on the right.
    b.edit("insert_picture", picture=Build.pick(all_paths, "/Purchase_Header/CompanyPicture"),
           locationType="documentEnd", widthMm=35, heightMm=18)
    b.note("logo PICTURE control bound to /Purchase_Header/CompanyPicture")

    addr = b.edit("insert_table", rowCount=6, columns=widths_and_alignments([4600, 4600]),
                  locationType="documentEnd")["data"]
    addr_idx = addr["tableIndex"]
    for row in range(6):
        b.edit("insert_field", field=f"/Purchase_Header/CompanyAddress{row + 1}",
               locationType="tableCell", tableIndex=addr_idx, row=row, col=0)
        b.edit("insert_field", field=f"/Purchase_Header/BuyFromAddr{row + 1}",
               locationType="tableCell", tableIndex=addr_idx, row=row, col=1)
    b.note("2x6 unbound table filled with CompanyAddress1-6 / BuyFromAddr1-6 FIELD controls (VendAddr1-6, "
           "which round 3 used, printed empty: the stock order layout prints the buy-from address)")

    # Document info: label/value pairs in a small borderless grid.
    info_tbl = b.edit("insert_table", rowCount=3, columns=widths_and_alignments([2800, 3000]),
                      locationType="documentEnd")["data"]
    info_idx = info_tbl["tableIndex"]
    pairs = [
        (Build.pick(all_paths, "/Purchase_Header/OrderNo_Lbl"),
         Build.pick(all_paths, "/Purchase_Header/No_PurchHeader")),
        (Build.pick(all_paths, "/Purchase_Header/DocumentDate_Lbl"),
         Build.pick(all_paths, "/Purchase_Header/DocumentDate")),
        (Build.pick(all_paths, "/Purchase_Header/PaymentTermsDesc_Lbl"),
         Build.pick(all_paths, "/Purchase_Header/PaymentTermsDesc")),
    ]
    for row, (label, field) in enumerate(pairs):
        b.edit("insert_label", label=label, locationType="tableCell",
               tableIndex=info_idx, row=row, col=0)
        b.edit("insert_field", field=field, locationType="tableCell",
               tableIndex=info_idx, row=row, col=1)
    b.note("3x2 label/value grid: order no, document date, payment terms (LABEL + FIELD controls)")

    # The typography knobs on top of the pinned default: a 16pt bold section title and a default-size
    # note, both plain static runs - the sizes are the ask, the text is just a carrier.
    b.edit("insert_text", text="ORDER LINES", locationType="documentEnd",
           format={"bold": True, "fontSizePoints": 16})
    b.edit("insert_text", text="All prices are in the order currency, excluding VAT.",
           locationType="documentEnd")
    b.note("insert_text twice: 'ORDER LINES' at 16pt bold + a default-size note (format bold/size "
           "on top of the pinned Calibri 11pt base)")

    # The line repeater, styled via the scaffolded TableGrid STYLE - a reference that resolved to
    # nothing before issue #3's fix.
    rep = b.edit("insert_repeater_table", dataItem="/Purchase_Header/Purchase_Line",
                 columns=repeater_columns(
                     "/Purchase_Header/Purchase_Line",
                     ["No_PurchLine", "Desc_PurchLine", "Qty_PurchLine", "UOM_PurchLine", "DirUnitCost_PurchLine",
                      "LineAmt_PurchLine"],
                     [1400, 3400, 900, 900, 1600, 1500], ["left", "left", "right", "left", "right", "right"]),
                 locationType="documentEnd", headerFromLabels=True, tableStyle="TableGrid")["data"]
    b.note(f"repeater table over /Purchase_Header/Purchase_Line, {rep['columnCount']} columns, "
           f"tableStyle='TableGrid' (the scaffolded style definition - w:tblStyle reference, no direct "
           f"border formatting)")

    # Totals: right-anchored block with a rule above it.
    totals = b.edit("insert_table", rowCount=1, columns=widths_and_alignments([6200, 2400], ["right", "right"]),
                    locationType="documentEnd")["data"]
    tot_idx = totals["tableIndex"]
    b.edit("insert_field", field=Build.pick(all_paths, "/Purchase_Header/Totals/TotalInclVATText"),
           locationType="tableCell", tableIndex=tot_idx, row=0, col=0)
    b.edit("insert_field", field=Build.pick(all_paths, "/Purchase_Header/Totals/TotalAmountInclVAT"),
           locationType="tableCell", tableIndex=tot_idx, row=0, col=1, format={"bold": True})
    b.edit("set_cell_format", tableIndex=tot_idx, row=0, edges=["top"], borderSize=4)
    b.note("right-anchored totals block with a 1/2pt rule above it (set_cell_format)")


@item(
    id="p07", title="Commission report from scratch incl scaffolded header and footer",
    report="115 - Salesperson Commission",
    base=None, schema="SalespersonCommission.docx",
    purpose="A third from-scratch build on an OLD-style report (a dedicated Labels data item, no "
            "*_Lbl sibling columns) that authors into the BLANK build's own scaffolded header/footer "
            "parts and rebuilds a nested detail repeater from nothing - P03 proved header/footer "
            "bindings on a stock layout's own parts; nobody has ever BC-rendered the parts "
            "create_layout scaffolds itself.",
    asks=[
        "TYPOGRAPHY (issue #3): same typeface and size in BC as in the mock (Calibri 11pt), including "
        "inside the header, the footer and the nested detail rows?",
        "The header and footer parts were CREATED by the tool's blank-build scaffold (empty part + "
        "sectPr reference), not by Word - does BC render them, with data, on every page?",
        "This report uses the dedicated Labels data item convention (no *_Lbl columns) - do the bound "
        "labels in the header/footer carry their caption text?",
        "The salesperson table's headers are humanized STATIC text (no matching label columns exist "
        "for headerFromLabels to bind) - acceptable, or visibly wrong against the data?",
        "Does the nested ledger-entry detail row, built from scratch with insert_repeater_row, repeat "
        "once per entry under each salesperson without drifting off the parent grid?",
    ],
    bc_steps=[
        "Search (Tell Me) for 'Salesperson Commission' (report 115) and set a posting-date filter wide "
        "enough to catch Cronus's posted sales (e.g. this fiscal year).",
        "Report Layouts > New layout for report 115, upload this .docx, then run the report with it - "
        "ideally to 2+ pages so the header/footer repetition is visible.",
    ],
)
def p07(b: Build):
    ds = b.dataset()
    all_paths = Build.paths(ds)

    # Header part: the report-name label, bound INLINE with a static separator + page marker. The part
    # itself is the blank build's scaffold - exercising it is this item's reason to exist.
    title_lbl = b.edit("insert_label", label=Build.pick(all_paths, "/Labels/ReportNameLabel",
                                                        "/Labels/SalespersonCommissionLabel"),
                       locationType="documentEnd", layoutPart="header")["data"]["controlId"]
    b.edit("insert_text", text="  |  scaffolded header", locationType="afterControl",
           controlId=title_lbl, layoutPart="header")
    b.note("scaffolded HEADER part: report-name LABEL control + static marker text, inline")

    # Footer part: the all-amounts-in-LCY note - a label column again, per the Labels-item convention.
    footer_lbl = b.edit("insert_label", label=Build.pick(all_paths, "/Labels/AllAmountsAreInLCYLabel",
                                                         "/Labels/RunOnLabel"),
                        locationType="documentEnd", layoutPart="footer")["data"]["controlId"]
    b.edit("insert_text", text="  |  scaffolded footer", locationType="afterControl",
           controlId=footer_lbl, layoutPart="footer")
    b.note("scaffolded FOOTER part: LCY-note LABEL control + static marker text, inline")

    # The salesperson repeater. No *_Lbl sibling columns exist on this report, so headerFromLabels
    # deliberately finds nothing and every header cell falls back to humanized static text.
    rep = b.edit("insert_repeater_table", dataItem="/Salesperson_Purchaser",
                 columns=repeater_columns(
                     "/Salesperson_Purchaser",
                     ["Salesperson_Purchaser_Code", "Salesperson_Purchaser_Name", "Salesperson_Purchaser__Commission___",
                      "SalesCommissionAmt", "AdjProfitCommissionAmt"],
                     [1400, 3306, 1500, 2000, 2000], ["left", "left", "right", "right", "right"]),
                 locationType="documentEnd", headerFromLabels=True)["data"]
    b.note(f"repeater table over /Salesperson_Purchaser, {rep['columnCount']} columns, humanized "
           f"static headers (this old-style report has no *_Lbl columns to bind)")

    # The nested detail: one row per customer ledger entry under each salesperson, built from nothing
    # (P04 rebuilt an EXISTING nested repeater; this one never existed).
    b.edit("insert_repeater_row", parentControlId=rep["controlId"],
           dataItem="/Salesperson_Purchaser/Cust_Ledger_Entry",
           cells=[{}] + [
               {"fields": [f"/Salesperson_Purchaser/Cust_Ledger_Entry/{name}"], "alignment": alignment}
               for name, alignment in [("Cust__Ledger_Entry__Posting_Date_", "left"),
                                       ("Cust__Ledger_Entry__Document_No__", "left"),
                                       ("CustLedgerEntry_CustomerName", "left"),
                                       ("Cust__Ledger_Entry__Sales__LCY__", "right")]])
    b.note("nested DETAIL ROW over /Salesperson_Purchaser/Cust_Ledger_Entry via insert_repeater_row: "
           "posting date, document no, customer, sales LCY - one row per entry, first column left "
           "as an indent spacer")

    # Grand totals: right-anchored, rule above.
    totals = b.edit("insert_table", rowCount=1, columns=widths_and_alignments([6200, 2400], ["right", "right"]),
                    locationType="documentEnd")["data"]
    tot_idx = totals["tableIndex"]
    b.edit("insert_label", label=Build.pick(all_paths, "/Labels/TotalsCaption"),
           locationType="tableCell", tableIndex=tot_idx, row=0, col=0)
    b.edit("insert_field", field=Build.pick(all_paths, "/Totals/Totals_SalesCommission"),
           locationType="tableCell", tableIndex=tot_idx, row=0, col=1, format={"bold": True})
    b.edit("set_cell_format", tableIndex=tot_idx, row=0, edges=["top"], borderSize=4)
    b.note("right-anchored totals block: TotalsCaption LABEL + total sales commission FIELD, "
           "1/2pt rule above")


# ---------------------------------------------------------------- research probes (round 3)
#
# P08-P11 do not check a capability: each settles an open issue that is blocked on BC evidence. A
# REJECTED upload is an answer too - record the exact message.

def rewrite_parts(path: str, pattern: str, transform) -> list[str]:
    """Rewrites every package part whose name matches `pattern` through `transform(xml) -> xml`, for the
    probes whose shape no tool emits (P09-P11, P18-P20). Returns the rewritten part names."""
    rewritten = []
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if re.fullmatch(pattern, info.filename):
                data = transform(data.decode("utf-8")).encode("utf-8")
                rewritten.append(info.filename)
            zout.writestr(info, data)
    # A virus scan can hold a freshly written .docx for a moment ("Access is denied"), as the server's own
    # commit allows for (TransientFileRetry).
    for attempt in range(5):
        try:
            os.replace(tmp, path)
            break
        except PermissionError:
            if attempt == 4:
                raise
            time.sleep(0.2 * (attempt + 1))
    return rewritten


@item(
    id="p08", title="Bindings whose storeItemID names no part",
    report="1306 - Standard Sales Invoice", base="StandardSalesInvoiceVatSpec.docx",
    purpose="GitHub issue #102, now a regression check. Microsoft's own StandardSalesInvoiceVatSpec ships two "
            "bindings (DocumentTitle_Lbl, DocumentNo) whose storeItemID names no part in the package. Round 3 "
            "printed them with their data, so validate_layout now reports that case as a warning, and "
            "refresh_xml_part re-points such bindings at the BC part. This item runs refresh_xml_part against "
            "the layout's own schema, which also repairs their foreign Mini_Sales_Invoice namespace (BC rejects "
            "that on upload), so both bindings name the BC part again.",
    asks=[
        "Does BC accept the upload, as it did in round 3?",
        "Do the document title (DocumentTitle_Lbl) and the invoice number (DocumentNo), top of page 1, still "
        "render WITH data now that refresh_xml_part re-pointed them?",
    ],
    bc_steps=[
        "Posted sales invoice: Sales > Posted Sales Invoices, pick any.",
        "Report Layouts > New layout for report 1306, upload this .docx, then Print the posted invoice "
        "with it.",
    ],
)
def p08(b: Build):
    refresh = b.edit("refresh_xml_part",
                     newSchemaSource=os.path.join(CORPUS_DIR, "StandardSalesInvoiceVatSpec.docx"))["data"]
    repointed = [f"{x['alias']} ({x['previousNamespace']})" for x in refresh["repointedForeignBindings"]]
    stores = [str(x["alias"]) for x in refresh["repointedStoreItemIds"]]
    if not stores:
        raise ItemFailure("refresh_xml_part re-pointed no storeItemID; the item no longer tests #102")
    b.note("refresh_xml_part against the layout's own schema: re-pointed the foreign-namespace bindings "
           f"{', '.join(repointed) or '(none)'}, and the storeItemIDs naming no part of {', '.join(stores)}")


# The add-in's two hide-if-empty wrappers differ only in their alias and tag: a cell-level w:sdt with no binding
# and w15:color FFA500, around a cell bound to a column of a CHILD repeating item (captured in
# SalespersonCommission and SubcontractorDispatchList). Round 3 ran them on report 115, which prints only
# salespeople with ledger entries, so neither wrapper had an empty case to act on. Round 3's re-run puts the same
# wrapper around the stock quote's caption cell bound to the FIRST line's ItemNo_Line_Lbl, which is empty exactly
# when the quote has no lines - a case any sandbox can produce.
HIDE_ROW_WRAPPER = ("Hide Table Row If Empty", "#BC:HideTableRowIfEmpty")
HIDE_TABLE_WRAPPER = ("Hide Table If Empty", "#BC:HideTableIfEmpty")
HIDDEN_CAPTION_XPATH = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line_Lbl[1]"


def element_end(xml: str, start: int, tag: str) -> int:
    """The index just past the w:<tag> element that opens at start, counting the ones nested in it."""
    depth = 0
    for m in re.compile(rf"<w:{tag}(?=[\s>/])[^>]*?(/)?>|</w:{tag}>").finditer(xml, start):
        if m.group(0).startswith("</"):
            depth -= 1
        elif m.group(1) != "/":
            depth += 1
        if depth == 0:
            return m.end()
    raise ItemFailure(f"unbalanced w:{tag} in document.xml")


def enclosing_start(xml: str, pos: int, tag: str) -> int:
    """Where the innermost w:<tag> element holding position pos opens."""
    for m in reversed(list(re.compile(rf"<w:{tag}(?=[\s>/])").finditer(xml, 0, pos))):
        if element_end(xml, m.start(), tag) > pos:
            return m.start()
    raise ItemFailure(f"no w:{tag} around the control")


def control_start(xml: str, xpath: str | None = None, control_id: int | None = None) -> int:
    """Where the one control bound to xpath, or carrying w:id control_id, opens."""
    if xpath is not None:
        binding = f'w:xpath="{xpath}"'
        if xml.count(binding) != 1:
            raise ItemFailure(f"expected one control bound to {xpath}, found {xml.count(binding)}")
        return xml.rfind("<w:sdt>", 0, xml.index(binding))
    ids = list(re.finditer(rf'<w:id w:val="{control_id}"\s*/>', xml))
    if len(ids) != 1:
        raise ItemFailure(f"expected one control with id {control_id}, found {len(ids)}")
    return xml.rfind("<w:sdt>", 0, ids[0].start())


def lift_inline_field_to_cell(xml: str, start: int) -> str:
    """The inline field opening at start, alone in its cell's one paragraph, becomes a CELL-level control around
    the w:tc - the shape the add-in's hide-if-empty wrapper encloses (no tool writes it)."""
    field = xml[start:element_end(xml, start, "sdt")]
    tc_start = enclosing_start(xml, start, "tc")
    tc_stop = element_end(xml, tc_start, "tc")
    p_start = enclosing_start(xml, start, "p")
    if p_start < tc_start:
        raise ItemFailure("the field is not in a table cell")
    paragraph = xml[p_start:element_end(xml, p_start, "p")]
    tc = xml[tc_start:tc_stop]
    tc_open = re.match(r"<w:tc\b[^>]*>", tc).group(0)
    tc_pr = (re.match(r"<w:tcPr>.*?</w:tcPr>|<w:tcPr/>", tc[len(tc_open):], re.S) or re.match("", "")).group(0)
    if tc[len(tc_open) + len(tc_pr):-len("</w:tc>")] != paragraph:
        raise ItemFailure("the cell holds more than the field's paragraph")
    p_open = re.match(r"<w:p\b[^>]*>", paragraph).group(0)
    p_pr = (re.match(r"<w:pPr>.*?</w:pPr>|<w:pPr/>", paragraph[len(p_open):], re.S) or re.match("", "")).group(0)
    if paragraph[len(p_open) + len(p_pr):-len("</w:p>")] != field:
        raise ItemFailure("the paragraph holds more than the field")
    parts = re.fullmatch(r"<w:sdt>(<w:sdtPr>.*?</w:sdtPr>)(<w:sdtEndPr>.*?</w:sdtEndPr>)?<w:sdtContent>(.*)"
                         r"</w:sdtContent></w:sdt>", field, re.S)
    if parts is None:
        raise ItemFailure("unexpected field shape")
    sdt_pr, end_pr, runs = parts.groups()
    cell_control = (f"<w:sdt>{sdt_pr}{end_pr or ''}<w:sdtContent>{tc_open}{tc_pr}{p_open}{p_pr}{runs}</w:p></w:tc>"
                    "</w:sdtContent></w:sdt>")
    return xml[:tc_start] + cell_control + xml[tc_stop:]


def wrap_in_hide_if_empty(b: Build, alias: str, tag: str, *, xpath: str | None = None,
                          control_id: int | None = None, lift: bool = False) -> int:
    """Wraps a cell-level control in the add-in's hide-if-empty control, in the captured shape: a cell-level w:sdt
    with no binding and w15:color FFA500, reusing the wrapped control's placeholder entry, with an id no control in
    the package has. lift first turns an inline field alone in its cell into that cell-level control. No tool
    writes either, so it is a hand edit. Returns the wrapper's id."""
    with zipfile.ZipFile(b.work) as package:
        xml_parts = "".join(package.read(n).decode("utf-8", "replace") for n in package.namelist()
                            if n.startswith("word/") and n.endswith(".xml"))
    used = set(re.findall(r'<w:id w:val="(-?\d+)"', xml_parts))
    wrapper_id = next(n for n in range(1_230_000_001, 1_230_100_000) if str(n) not in used)

    def wrap(xml: str) -> str:
        if lift:
            xml = lift_inline_field_to_cell(xml, control_start(xml, xpath, control_id))
        start = control_start(xml, xpath, control_id)
        end = element_end(xml, start, "sdt")
        control = xml[start:end]
        if "<w:sdtContent><w:tc>" not in control and not re.search(r"<w:sdtContent><w:tc\b", control):
            raise ItemFailure("the wrapped control is not a cell-level control")
        doc_part = re.search(r'<w:docPart w:val="([^"]+)"\s*/>', control).group(1)
        wrapper = (f'<w:sdt><w:sdtPr><w:alias w:val="{alias}"/><w:tag w:val="{tag}"/><w:id w:val="{wrapper_id}"/>'
                   f'<w:placeholder><w:docPart w:val="{doc_part}"/></w:placeholder><w15:color w:val="FFA500"/>'
                   f"</w:sdtPr><w:sdtContent>{control}</w:sdtContent></w:sdt>")
        return xml[:start] + wrapper + xml[end:]

    rewrite_parts(b.work, r"word/document\.xml", wrap)
    return wrapper_id


def wrap_caption_in_hide_if_empty(b: Build, alias: str, tag: str):
    """P09/P10: the quote's ItemNo_Line_Lbl caption cell, already a cell-level control, in the wrapper."""
    wrapper_id = wrap_in_hide_if_empty(b, alias, tag, xpath=HIDDEN_CAPTION_XPATH)
    b.note(f"hand edit in word/document.xml: the lines table's caption cell bound to /Header/Line/ItemNo_Line_Lbl "
           f"wrapped in a '{alias}' control ({tag}, id {wrapper_id}), in the shape the add-in writes - as captured "
           "in Salesperson Commission")


@item(
    id="p09", title="Hide Table Row If Empty on a quote's caption row",
    report="1304 - Standard Sales Quote", base="StandardSalesQuote.docx",
    purpose="GitHub issue #123: the BC add-in's 'Hide Table Row If Empty' control is captured (a cell-level "
            "wrapper tagged #BC:HideTableRowIfEmpty, in Salesperson Commission) but not what BC treats as EMPTY. "
            "Round 3 ran the stock commission layout, and report 115 prints only salespeople with ledger entries, "
            "so the empty case never came up. Here the same wrapper, hand-made in the captured shape, goes around "
            "the stock quote's caption cell bound to /Header/Line/ItemNo_Line_Lbl: a column of the FIRST quote "
            "line, as the captured one binds a column of the first ledger entry, so it is empty exactly when the "
            "quote has no lines.",
    asks=[
        "A quote with NO lines: is the lines table's caption row (No., Description, Quantity, ...) left out? Is "
        "anything else missing, such as the totals rows or the whole table?",
        "A quote WITH lines: is the caption row printed normally, 'No.' included, and nothing else hidden?",
    ],
    bc_steps=[
        "Two quotes: the one P01/P05/P13 use (several lines), and a new quote for any customer with NO lines.",
        "Report Layouts > New layout for report 1304, upload this .docx, then Print both quotes with it.",
    ],
)
def p09(b: Build):
    b.drop_broken_abn_control()
    wrap_caption_in_hide_if_empty(b, *HIDE_ROW_WRAPPER)


@item(
    id="p10", title="Hide Table If Empty on a quote's lines table",
    report="1304 - Standard Sales Quote", base="StandardSalesQuote.docx",
    purpose="GitHub issue #123, the table variant: 'Hide Table If Empty' (#BC:HideTableIfEmpty) is captured in the "
            "Subcontractor Dispatch List, shaped exactly like P09's row variant but for its alias and tag. This is "
            "P09's quote with that wrapper instead, around the same caption cell, so the renders of the same two "
            "quotes side by side show what each variant hides.",
    asks=[
        "A quote with NO lines: what is left out - the caption row, the whole lines table including its totals "
        "rows, or nothing?",
        "A quote WITH lines: is the table printed normally?",
        "Side by side with P09's render of the same two quotes: what does the table variant hide that the row "
        "variant does not?",
    ],
    bc_steps=[
        "The same two quotes as P09: one with several lines, one with NO lines.",
        "Report Layouts > New layout for report 1304, upload this .docx, then Print both quotes with it.",
    ],
)
def p10(b: Build):
    b.drop_broken_abn_control()
    wrap_caption_in_hide_if_empty(b, *HIDE_TABLE_WRAPPER)


# The nested IF PAGE = NUMPAGES construct, run for run as Word itself serializes it: opening this exact
# paragraph in Word for Microsoft 365 and saving it back (2026-10-17) left every run unchanged. A nested
# field's cached result sits inside the outer field's instruction, so it is a w:instrText, not a w:t.
LAST_PAGE_TEXT = "LAST PAGE ONLY"


def _fld(kind: str) -> str:
    return f'<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>'


def _instr(text: str) -> str:
    return f'<w:r><w:instrText xml:space="preserve">{text}</w:instrText></w:r>'


def _nested(instruction: str) -> str:
    return (_fld("begin") + _instr(instruction) + _fld("separate")
            + "<w:r><w:instrText>1</w:instrText></w:r>" + _fld("end"))


LAST_PAGE_PARAGRAPH = (
    "<w:p>"
    + _fld("begin") + _instr(" IF ") + _nested(" PAGE ") + _instr(" = ") + _nested(" NUMPAGES ")
    + _instr(f' "{LAST_PAGE_TEXT}" "" ')
    + _fld("separate") + f"<w:r><w:t>{LAST_PAGE_TEXT}</w:t></w:r>" + _fld("end")
    + "</w:p>"
)


@item(
    id="p11", title="Last-page-only footer text via IF PAGE = NUMPAGES",
    report="1322 - Standard Purchase Order", base="StandardPurchaseOrder.docx",
    purpose="GitHub issue #11: 'company name in the footer on the last page only' is a Word IF PAGE = "
            "NUMPAGES field. No tool emits it and no corpus layout carries one, so this item adds it by "
            "hand: a paragraph reading LAST PAGE ONLY, in every footer part, inside { IF { PAGE } = "
            "{ NUMPAGES } \"LAST PAGE ONLY\" \"\" }. In Word (the mock preview) it shows on the last page "
            "only; whether BC's renderer evaluates IF/PAGE/NUMPAGES per page is the open question.",
    asks=[
        "Does BC accept the upload?",
        "A MULTI-page order: does LAST PAGE ONLY appear on the last page and on no other?",
        "A ONE-page order: does it appear (page 1 is the last page)?",
        "If it shows on every page or none: which, and does the stock 'Page X / Y' in the header still "
        "count correctly on that render?",
    ],
    bc_steps=[
        "Purchase order with enough lines for 3+ pages (the mock needed ~30 lines); add lines to an open "
        "Cronus order if none is that long. Also pick a one-line order.",
        "Report Layouts > New layout for report 1322, upload this .docx, then Print both orders with it.",
    ],
)
def p11(b: Build):
    parts = rewrite_parts(b.work, r"word/footer\d*\.xml",
                          lambda xml: xml.replace("</w:ftr>", LAST_PAGE_PARAGRAPH + "</w:ftr>", 1))
    b.note(f"hand-edited (no tool emits it): IF PAGE = NUMPAGES paragraph appended to {', '.join(parts)}")


# ---------------------------------------------------------------- what 2.0 writes (round 3)
#
# P12-P14 are capability checks again: each renders shapes the 2.0 tools write that no sandbox round has
# seen - a tool-written block repeater (P12, and the nested-item one in P14), the formatting tools (P13)
# and the reworked table tools (P14).

@item(
    id="p12", title="Statement repeated per customer by wrap_in_repeater",
    report="1316 - Standard Statement", base=None, schema="StandardStatement.docx",
    purpose="Issue #31 added wrap_in_repeater: it wraps body content in a block-level repeater, over a "
            "top-level item in the shape the stock statement carries (section > w:group > item; without the "
            "group BC prints an empty document). This statement is built from scratch for ONE customer - "
            "addresses, a statement-info grid, the detailed ledger entries - and then the whole body is "
            "wrapped over /Customer with a page break before every customer. No tool-written block repeater "
            "has been rendered by BC yet.",
    asks=[
        "Does BC accept the upload and print the statements - not an empty document?",
        "A statement for two or more customers: does every customer start on a new page, with its OWN "
        "addresses, dates and entries?",
        "Is there a blank page before the first customer? (The page break sits before every customer, the "
        "first included; Word's mock shows none.)",
        "The entries table lists the detailed entries of the customer's FIRST currency only (this layout "
        "has no currency loop): do they match that customer's entries in BC?",
    ],
    bc_steps=[
        "Search (Tell Me) for 'Statement' (report 1316). Pick two or three Cronus customers with posted "
        "entries (e.g. 10000..30000) and a date range that covers them.",
        "Report Layouts > New layout for report 1316, upload this .docx, then run the statement for those "
        "customers with it, printed to ONE PDF.",
    ],
)
def p12(b: Build):
    integer = "/Customer/Integer"
    addr = b.edit("insert_table", rowCount=5, columns=widths_and_alignments([4600, 4600]),
                  locationType="documentEnd")["data"]
    for row in range(5):
        b.edit("insert_field", field=f"{integer}/CompanyAddr{row + 1}", locationType="tableCell",
               tableIndex=addr["tableIndex"], row=row, col=0)
        b.edit("insert_field", field=f"{integer}/CustAddr{row + 1}", locationType="tableCell",
               tableIndex=addr["tableIndex"], row=row, col=1)
    b.note("2x5 address table: CompanyAddr1-5 beside CustAddr1-5, all under /Customer/Integer")

    info = b.edit("insert_table", rowCount=3, columns=widths_and_alignments([2800, 3000]),
                  locationType="documentEnd")["data"]
    for row, (caption, value) in enumerate([("StartDateCaption", "StartDate"), ("EndDateCaption", "EndDate"),
                                            ("No1_CustCaption", "No1_Cust")]):
        b.edit("insert_field", field=f"{integer}/{caption}", locationType="tableCell",
               tableIndex=info["tableIndex"], row=row, col=0, format={"bold": True})
        b.edit("insert_field", field=f"{integer}/{value}", locationType="tableCell",
               tableIndex=info["tableIndex"], row=row, col=1)
    b.note("3x2 info grid: start date, end date, customer no., each caption a bound *Caption column")

    entries = f"{integer}/CurrencyLoop/CustLedgEntryHdr/DtldCustLedgEntries"
    columns = [("PostDate_DtldCustLedgEntries", 1300, None, "PostDate_DtldCustLedgEntriesCaption"),
               ("DocNo_DtldCustLedgEntries", 1600, None, "DocNo_DtldCustLedgEntriesCaption"),
               ("Description", None, None, "Desc_CustLedgEntry2Caption"),
               ("DueDate_DtldCustLedgEntries", 1300, None, "DueDate_CustLedgEntry2Caption"),
               ("Amt_DtldCustLedgEntries", 1500, "right", "OriginalAmt_CustLedgEntry2Caption"),
               ("RemainAmt_DtldCustLedgEntries", 1500, "right", "RemainAmtCustLedgEntry2Caption")]
    lines = b.edit("insert_repeater_table", dataItem=entries, locationType="documentEnd",
                   columns=[{k: v for k, v in (("field", f"{entries}/{name}"), ("width", width),
                                                ("alignment", alignment)) if v is not None}
                            for name, width, alignment, _ in columns])["data"]
    # The report's captions are *Caption columns of /Customer/Integer, not *_Lbl siblings of the entry
    # columns, so headerFromLabels finds none: bind them into the header row by hand.
    for col, (_, _, _, caption) in enumerate(columns):
        b.edit("clear_cell_text", tableIndex=lines["tableIndex"], row=0, col=col)
        b.edit("insert_field", field=f"{integer}/{caption}", locationType="tableCell",
               tableIndex=lines["tableIndex"], row=0, col=col, format={"bold": True})
    b.note(f"repeater table over {entries}: posting date, document no., description (no width: it takes "
           "what the others leave), due date, amount, remaining amount; header cells bound to the report's "
           "*Caption columns")

    wrapped = b.edit("wrap_in_repeater", dataItem="/Customer", wholeBody=True, pageBreakBefore=True)["data"]
    b.note(f"wrap_in_repeater over /Customer, wholeBody=true, pageBreakBefore=true: block repeater "
           f"{wrapped['controlId']} with its w:group around the whole body")


MISSING_FONT = "BCWL Missing Font"


@item(
    id="p13", title="Text and cell formatting written by the tools",
    report="1304 - Standard Sales Quote", base="StandardSalesQuote.docx",
    purpose="Issue #15 added a format object (bold, italic, underline, size, font, colour, character "
            "style) on every text tool, set_text_format to restyle existing text, and a cell fill on "
            "set_cell_format. Each writes a shape the stock layouts carry - except italic, which the corpus "
            "applies only through style definitions, never directly on a run.",
    asks=[
        "ITALIC: the line 'P13 italic (direct w:i)' carries a direct w:i + w:iCs, a shape no stock layout "
        "has. Is it italic in BC, as in the mock?",
        "Do 'P13 Segoe UI 10 pt bold blue', 'P13 underlined' and 'P13 Strong character style' look as "
        "in the mock (blue is 0070C0, the stock accent)?",
        f"FONTS: 'P13 Georgia' in Georgia, and 'P13 {MISSING_FONT}' - a font that exists nowhere - in a "
        "substitute (round 3: a serif face)? validate_layout now warns about the second (font-installed), "
        "and insert_text's summary says so.",
        "RESTYLED STOCK CONTENT: is the document number in the page header blue and italic "
        "(set_text_format on the control), and is the line table's header row shaded light grey with "
        "underlined captions (set_cell_format fill + set_text_format on the row)?",
        "In 'P13 only THIS WORD is red', is only THIS WORD red (set_text_format split it out of its run)?",
    ],
    bc_steps=[
        "Use the SAME sales quote as P01/P05 so the renders are directly comparable.",
        "Report Layouts > New layout for report 1304, upload, then Print the quote with it. The test lines "
        "are at the end of the body, after the totals.",
    ],
)
def p13(b: Build):
    b.drop_broken_abn_control()

    # Restyle stock content that is already there.
    info = b.info()
    # The quote carries DocumentNo in its default and its first-page header; restyle both, so page 1
    # shows it whichever part BC renders there.
    doc_nos = [c for c in info["controls"]
               if str(c.get("alias") or "").endswith("/Header/DocumentNo") and c["part"].startswith("header")]
    if not doc_nos:
        raise ItemFailure("no DocumentNo control in a header part")
    for doc_no in doc_nos:
        b.edit("set_text_format", target="control", controlId=doc_no["sdtId"],
               format={"color": "0070C0", "italic": True})
    b.note("set_text_format on the DocumentNo control in "
           f"{', '.join(c['part'] for c in doc_nos)}: blue 0070C0, italic, over its stock 'Strong' style")

    lines = Build.line_table(info, Build.repeater(info, "/Header/Line"))
    b.edit("set_cell_format", tableIndex=lines, row=0, fill="E8E8E8")
    b.edit("set_text_format", target="row", tableIndex=lines, row=0, format={"underline": True})
    b.note(f"line table {lines} header row: set_cell_format fill E8E8E8 + set_text_format underline")

    # One line per format member, at the end of the body.
    samples = [
        ("P13 Segoe UI 10 pt bold blue", {"fontName": "Segoe UI", "fontSizePoints": 10, "bold": True,
                                          "color": "0070C0"}),
        ("P13 italic (direct w:i)", {"italic": True}),
        ("P13 underlined", {"underline": True}),
        ("P13 Strong character style", {"characterStyle": "Strong"}),
        ("P13 Georgia", {"fontName": "Georgia"}),
        (f"P13 {MISSING_FONT}", {"fontName": MISSING_FONT}),
    ]
    for text, fmt in samples:
        b.edit("insert_text", text=text, locationType="documentEnd", format=fmt)
    b.note("insert_text x6 at the end of the body, one format member each: Segoe UI 10 pt bold blue, "
           f"direct italic, underline, the 'Strong' character style, Georgia, and '{MISSING_FONT}'")

    b.edit("insert_text", text="P13 only THIS WORD is red", locationType="documentEnd")
    b.edit("set_text_format", target="text", searchText="THIS WORD", format={"color": "C00000"})
    b.note("set_text_format text 'THIS WORD': colour C00000 on the matched words only (run split)")


@item(
    id="p14", title="Quote with the 2.0 table options",
    report="1304 - Standard Sales Quote", base=None, schema="StandardSalesQuote.docx",
    purpose="2.0 reworked the table tools (issues #131-#133): plain tables work in headers and footers, "
            "both table tools take one borders option (none / headerRule / grid), and a column without a "
            "width shares what the others leave. This quote is built from scratch with them: the logo beside "
            "the title and number in a header table, the company details in a footer table under a rule, a "
            "caption-over-value info grid with a header rule, a lines table with a rule under every line, a "
            "gridded totals block, and the work description repeated per line by wrap_in_repeater (issue "
            "#31, the nested-item shape). The first re-run printed no work description although the quote had "
            "one: the repeater held the field INLINE in a paragraph, a shape no stock layout has. wrap_in_repeater "
            "now writes it as a block-level control around the paragraph, the stock quote's shape (P18 checks "
            "the old shape on the stock layout).",
    asks=[
        "HEADER AND FOOTER TABLES: on every page, is the logo beside 'Sales Quote <number>' at the top, and "
        "are the three company-address lines at the bottom, under a thin rule?",
        "INFO GRID: bold 8 pt captions over their values, with a thin rule under the caption row only "
        "(borders='headerRule' on a plain table)?",
        "LINES: does EVERY line carry a thin rule under it (the rule is on the repeated row), and the "
        "header row its own? Is the description the widest column (it has no width of its own)?",
        "TOTALS: is the totals block gridded on every edge and between its cells (borders='grid' writes "
        "explicit table borders, which no sandbox round has rendered)?",
        "WORK DESCRIPTION: with a work description of two or more lines, does each line print as its own "
        "paragraph, in order, below the totals? (The first re-run printed none.)",
    ],
    bc_steps=[
        "Use the SAME sales quote as P01/P05/P13, and give it a Work Description of two or three lines "
        "first (the field on the quote card).",
        "Report Layouts > New layout for report 1304, upload this .docx, then Print the quote with it.",
    ],
)
def p14(b: Build):
    all_paths = Build.paths(b.dataset())

    header = b.edit("insert_table", rowCount=1, columns=widths_and_alignments([3600, None], [None, "right"]),
                    locationType="documentEnd", layoutPart="header")["data"]
    in_header = {"locationType": "tableCell", "tableIndex": header["tableIndex"], "row": 0, "layoutPart": "header"}
    b.edit("insert_picture", picture=Build.pick(all_paths, "/Header/CompanyPicture"), col=0,
           widthMm=35, heightMm=18, **in_header)
    title = b.edit("insert_label", label=Build.pick(all_paths, "/Header/DocumentTitle_Lbl"), col=1,
                   format={"bold": True, "fontSizePoints": 14}, **in_header)["data"]
    b.edit("insert_field", field=Build.pick(all_paths, "/Header/DocumentNo"), locationType="afterControl",
           controlId=title["controlId"], layoutPart="header")
    b.edit("insert_text", text=" ", locationType="afterControl", controlId=title["controlId"],
           layoutPart="header")
    b.note(f"HEADER table ({header['part']}, borderless, 3600 + the rest): logo picture control beside the "
           "DocumentTitle_Lbl label, a space and DocumentNo, right-aligned, bold 14 pt")

    footer = b.edit("insert_table", rowCount=1, columns=widths_and_alignments([None, None, None]),
                    locationType="documentEnd", layoutPart="footer")["data"]
    for col in range(3):
        b.edit("insert_field", field=f"/Header/CompanyAddress{col + 1}", locationType="tableCell",
               tableIndex=footer["tableIndex"], row=0, col=col, layoutPart="footer", format={"fontSizePoints": 8})
    b.edit("set_cell_format", tableIndex=footer["tableIndex"], row=0, edges=["top"], layoutPart="footer")
    b.note(f"FOOTER table ({footer['part']}, three even columns): CompanyAddress1-3 at 8 pt, a 1/2 pt "
           "rule above (set_cell_format top)")

    info = b.edit("insert_table", rowCount=2, columns=widths_and_alignments([2600, 2600, 2600]),
                  locationType="documentEnd", borders="headerRule")["data"]
    pairs = [
        (Build.pick(all_paths, "/Header/DocumentDate_Lbl"), Build.pick(all_paths, "/Header/DocumentDate")),
        (Build.pick(all_paths, "/Header/QuoteValidToDate_Lbl"), Build.pick(all_paths, "/Header/QuoteValidToDate")),
        (Build.pick(all_paths, "/Header/SelltoCustomerNo_Lbl", "/Header/BilltoCustomerNo_Lbl"),
         Build.pick(all_paths, "/Header/SelltoCustomerNo", "/Header/BilltoCustumerNo")),
    ]
    for col, (label, field) in enumerate(pairs):
        b.edit("insert_label", label=label, locationType="tableCell", tableIndex=info["tableIndex"], row=0,
               col=col, format={"bold": True, "fontSizePoints": 8})
        b.edit("insert_field", field=field, locationType="tableCell", tableIndex=info["tableIndex"], row=1,
               col=col)
    b.note("2x3 info grid, borders='headerRule': bold 8 pt captions over document date, valid-to date and "
           "customer no.")

    lines = b.edit("insert_repeater_table", dataItem="/Header/Line", locationType="documentEnd",
                   headerFromLabels=True,
                   columns=[{"field": "/Header/Line/ItemNo_Line", "width": 1500},
                            {"field": "/Header/Line/Description_Line"},
                            {"field": "/Header/Line/Quantity_Line", "width": 1100, "alignment": "right"},
                            {"field": "/Header/Line/UnitPrice", "width": 1400, "alignment": "right"},
                            {"field": "/Header/Line/LineAmount_Line", "width": 1600, "alignment": "right"}])["data"]
    b.edit("set_cell_format", tableIndex=lines["tableIndex"], row=lines["dataRowIndex"], edges=["bottom"])
    b.note("lines table over /Header/Line (borders default 'headerRule'; the description has no width), "
           "plus set_cell_format bottom on the repeating data row: a rule under every line")

    totals = b.edit("insert_table", rowCount=2, columns=widths_and_alignments([6200, 2400], ["right", "right"]),
                    locationType="documentEnd", borders="grid")["data"]
    in_totals = {"locationType": "tableCell", "tableIndex": totals["tableIndex"]}
    b.edit("insert_field", field="/Header/Totals/TotalExcludingVATText", row=0, col=0, **in_totals)
    b.edit("insert_field", field="/Header/Totals/TotalNetAmount", row=0, col=1, **in_totals)
    b.edit("insert_field", field="/Header/Totals/TotalIncludingVATText", row=1, col=0, **in_totals)
    b.edit("insert_field", field="/Header/Totals/TotalAmountIncludingVAT", row=1, col=1, format={"bold": True},
           **in_totals)
    b.note("2x2 totals block, borders='grid': net and gross totals, the gross amount bold")

    line = b.edit("insert_field", field="/Header/WorkDescriptionLines/WorkDescriptionLine",
                  locationType="documentEnd")["data"]
    wrapped = b.edit("wrap_in_repeater", dataItem="/Header/WorkDescriptionLines",
                     firstControlId=line["controlId"])["data"]
    b.note(f"work-description line field in its own paragraph, wrapped by wrap_in_repeater over "
           f"/Header/WorkDescriptionLines (block repeater {wrapped['controlId']}, no group: a nested item)")



# P15-P17 put what no round has rendered from the tools at all in front of BC: 1.1.0's totals rows inside a
# lines table, page numbers and group subtotals - stock shapes, never printed from tool output - and
# create_layout over a template, which 2.0 changed (#77).

@item(
    id="p15", title="Quote totals inside the lines table, with page numbers",
    report="1304 - Standard Sales Quote", base=None, schema="StandardSalesQuote.docx",
    purpose="Two 1.1.0 tools emit shapes copied from Microsoft's layouts that no sandbox round has printed: "
            "insert_table_row, the static totals rows every stock document puts INSIDE its lines table, and "
            "insert_page_number, the stock 'Page X / Y' field codes. This quote is built from scratch with both: "
            "a lines table closed by a spacer row and a net / VAT / gross ladder (gross bold, under a rule), "
            "'Page X / Y' in the header and the bare page number in the footer.",
    asks=[
        "The totals rows belong to the lines table, after its repeating row: do they print exactly ONCE, after "
        "the last line - never between lines?",
        "A quote long enough for two pages: does the table's header row repeat on page 2, and do the totals "
        "print at the end only?",
        "Header: 'Page  1 / 2', then 'Page  2 / 2' (caption, two spaces, page / pages). Footer: 'Page 1', then "
        "'Page 2' (includeTotal=false). Do they count correctly on every page?",
        "A one-page quote: 'Page  1 / 1' and 'Page 1'?",
    ],
    bc_steps=[
        "Two quotes: the one-page quote P01/P05/P13/P14 use, and one with enough lines for two pages (about 30 "
        "short lines; add lines to an open Cronus quote if none is that long).",
        "Report Layouts > New layout for report 1304, upload this .docx, then Print both quotes with it.",
    ],
)
def p15(b: Build):
    for part, include_total, separator in (("header", True, "  "), ("footer", False, " ")):
        caption = b.edit("insert_label", label="/Header/Page_Lbl", locationType="documentEnd",
                         layoutPart=part)["data"]["controlId"]
        b.edit("insert_page_number", locationType="afterControl", controlId=caption, includeTotal=include_total,
               layoutPart=part)
        b.edit("insert_text", text=separator, locationType="afterControl", controlId=caption, layoutPart=part)
    b.note("HEADER: Page_Lbl, two spaces, PAGE / NUMPAGES (insert_page_number); FOOTER: Page_Lbl, a space and "
           "PAGE alone (includeTotal=false) - each built right to left on the label")

    lines = b.edit("insert_repeater_table", dataItem="/Header/Line", locationType="documentEnd",
                   headerFromLabels=True,
                   columns=[{"field": "/Header/Line/ItemNo_Line", "width": 1500},
                            {"field": "/Header/Line/Description_Line"},
                            {"field": "/Header/Line/Quantity_Line", "width": 1100, "alignment": "right"},
                            {"field": "/Header/Line/UnitPrice", "width": 1400, "alignment": "right"},
                            {"field": "/Header/Line/LineAmount_Line", "width": 1600, "alignment": "right"}])["data"]
    table = lines["tableIndex"]
    b.edit("insert_table_row", tableIndex=table, cells=[{"span": 5}])
    ladder = [("TotalExcludingVATText", "TotalNetAmount"), ("TotalVATAmountText", "TotalVATAmount"),
              ("TotalIncludingVATText", "TotalAmountIncludingVAT")]
    for caption, amount in ladder:
        b.edit("insert_table_row", tableIndex=table,
               cells=[{"span": 2}, {"fields": [f"/Header/Totals/{caption}"], "span": 2},
                      {"fields": [f"/Header/Totals/{amount}"], "alignment": "right"}],
               **({"format": {"bold": True}} if amount == "TotalAmountIncludingVAT" else {}))
    last_row = next(t for t in b.info()["tables"]
                    if t["part"] == "document.xml" and t["tableIndex"] == table)["rowCount"] - 1
    b.edit("set_cell_format", tableIndex=table, row=last_row, edges=["top"])
    b.note(f"lines table over /Header/Line closed by insert_table_row x4 INSIDE it: a spacer row, net, VAT and "
           f"gross totals (text field + amount), the gross row bold with a 1/2 pt rule above (row {last_row})")


@item(
    id="p16", title="Commission report with per-salesperson subtotals",
    report="115 - Salesperson Commission", base=None, schema="SalespersonCommission.docx",
    purpose="insert_subtotal_row (1.1.0) appends static rows at the END of a repeater's item - the stock "
            "group-subtotal shape of this very report: a spacer row, then a bold row bound to the sibling "
            "Subtotals item, once per salesperson after their ledger entries. No sandbox round has printed it "
            "from tool output. This is P07's salesperson table with its nested entry rows, closed per group by "
            "those two rows and, for the whole table, by a grand-total row under a rule.",
    asks=[
        "Each salesperson: their row (code, name, commission %, amounts), their ledger-entry rows, then an "
        "empty spacer row and a bold subtotal row. Does the subtotal pair print ONCE per salesperson, after "
        "their last entry - never per entry?",
        "Do the subtotal amounts match that salesperson's entries (they come from the report's Subtotals item)?",
        "Does the bold grand-total row print once, at the very end, under a rule, and match the sum of the "
        "subtotals?",
    ],
    bc_steps=[
        "Search (Tell Me) for 'Salesperson Commission' (report 115) with a posting-date filter that catches "
        "Cronus's posted sales for at least two salespeople (e.g. this fiscal year).",
        "Report Layouts > New layout for report 115, upload this .docx, then run the report with it.",
    ],
)
def p16(b: Build):
    group = "/Salesperson_Purchaser"
    rep = b.edit("insert_repeater_table", dataItem=group, locationType="documentEnd", headerFromLabels=True,
                 columns=repeater_columns(
                     group, ["Salesperson_Purchaser_Code", "Salesperson_Purchaser_Name",
                             "Salesperson_Purchaser__Commission___", "SalesCommissionAmt", "AdjProfitCommissionAmt"],
                     [1400, 3306, 1500, 2000, 2000], ["left", "left", "right", "right", "right"]))["data"]
    b.edit("insert_repeater_row", parentControlId=rep["controlId"], dataItem=f"{group}/Cust_Ledger_Entry",
           cells=[{}] + [{"fields": [f"{group}/Cust_Ledger_Entry/{name}"], "alignment": alignment}
                         for name, alignment in [("Cust__Ledger_Entry__Posting_Date_", "left"),
                                                 ("Cust__Ledger_Entry__Document_No__", "left"),
                                                 ("SalesCommissionAmt_Control32", "right"),
                                                 ("AdjProfitCommissionAmt_Control45", "right")]])
    b.note(f"P07's shape: repeater over {group} (code, name, commission %, amounts) with a nested detail row "
           "per customer ledger entry (posting date, document no., its two commission amounts)")

    subtotals = f"{group}/Subtotals"
    b.edit("insert_subtotal_row", parentControlId=rep["controlId"], cells=[{"span": 5}])
    b.edit("insert_subtotal_row", parentControlId=rep["controlId"], format={"bold": True},
           cells=[{"fields": [f"{subtotals}/Subtotals_Salesperson_Code"]},
                  {"fields": ["/Labels/TotalsCaption"], "span": 2},
                  {"fields": [f"{subtotals}/Subtotals_SalesCommission"], "alignment": "right"},
                  {"fields": [f"{subtotals}/Subtotals_AdjProfitCommission"], "alignment": "right"}])
    b.note(f"insert_subtotal_row x2 at the end of each {group} item: a spacer row, then a bold row with the "
           f"salesperson code, the TotalsCaption label and {subtotals}' two commission totals")

    b.edit("insert_table_row", tableIndex=rep["tableIndex"], format={"bold": True},
           cells=[{"fields": ["/Labels/TotalsCaption"], "span": 3},
                  {"fields": ["/Totals/Totals_SalesCommission"], "alignment": "right"},
                  {"fields": ["/Totals/Totals_AdjProfitCommission"], "alignment": "right"}])
    last_row = next(t for t in b.info()["tables"]
                    if t["part"] == "document.xml" and t["tableIndex"] == rep["tableIndex"])["rowCount"] - 1
    b.edit("set_cell_format", tableIndex=rep["tableIndex"], row=last_row, edges=["top"])
    b.note("insert_table_row: a static bold grand-total row (TotalsCaption label, the /Totals commission "
           "totals) after the repeater, with a 1/2 pt rule above")


DOCUMENT_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
TEMPLATE_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"


def p17_template(b: Build, out_dir: str) -> str:
    """A branded .dotx made from the stock quote the way a designer would: the body emptied (its page setup and
    distinct first page kept), every bound control stripped out of the headers and footers, a blue marker line
    added to each header and footer part, and the package typed as a Word template."""
    shell = os.path.join(out_dir, "template-source.docx")
    shutil.copy2(os.path.join(CORPUS_DIR, "StandardSalesQuote.docx"), shell)

    def empty_body(xml: str) -> str:
        start = xml.index("<w:body>") + len("<w:body>")
        return xml[:start] + xml[xml.rindex("<w:sectPr"):]

    rewrite_parts(shell, r"word/document\.xml", empty_body)
    removed = 0
    while True:
        info = b.tool("get_layout_info", {"layoutPath": shell})["data"]
        bound = [c for c in info["controls"] if c["kind"] != "Unbound"]
        if not bound:
            break
        b.tool("remove_control", {"layoutPath": shell, "controlId": bound[0]["sdtId"]})
        removed += 1
    for part in info["partDetails"]:
        if part["kind"] in ("header", "footer") and part.get("role") in ("default", "first"):
            b.tool("insert_text", {"layoutPath": shell, "text": f"P17 TEMPLATE {part['kind'].upper()}",
                                   "locationType": "documentEnd", "layoutPart": part["kind"],
                                   "partName": part["name"], "format": {"bold": True, "color": "0070C0"}})

    template = os.path.join(out_dir, "template.dotx")
    shutil.move(shell, template)
    typed = rewrite_parts(template, r"\[Content_Types\]\.xml", lambda xml: xml.replace(DOCUMENT_MAIN, TEMPLATE_MAIN))
    if not typed:
        raise ItemFailure("the template's [Content_Types].xml was not rewritten")
    b.note(f"template.dotx, made from the stock quote: body emptied (hand edit; page setup and distinct first "
           f"page kept), {removed} bound controls removed from its headers and footers (remove_control), a "
           "bold blue 'P17 TEMPLATE HEADER/FOOTER' line in each default and first-page part (insert_text), "
           "and the main part typed as a Word template (hand edit). Its settings still link the stock quote's "
           "attached template, which create_layout removes (BC refuses to upload a layout linking one)")
    return template


@item(
    id="p17", title="Quote created over a branded .dotx template",
    report="1304 - Standard Sales Quote", base=None, schema="StandardSalesQuote.docx", template=p17_template,
    purpose="create_layout's templatePath builds a layout over an unbound branded shell, keeping its headers, "
            "footers, styles and settings. The from-design skill relies on that route, but only mock renders have "
            "checked it, and 2.0 changed it: a .dotx template is now converted to a .docx (#77), where it used to "
            "keep the template content type Word refuses to open. The template here is made from the stock quote "
            "(see template.dotx in this folder); create_layout builds over it, then the tools add the document "
            "title and number to both of its headers and author a plain body.",
    asks=[
        "Does BC accept a layout create_layout built over a .dotx template?",
        "Is the template's look kept: the blue 'P17 TEMPLATE HEADER' and 'P17 TEMPLATE FOOTER' lines on every "
        "page (page 1 from the template's first-page parts, as its distinct first page says), and the page "
        "numbers left in the stock header table counting correctly?",
        "Does the document title and number the tools added show in the header on page 1 AND on page 2 (one in "
        "each of the template's two headers)?",
        "Does the body the tools added match the mock's typeface and sizes (the template's own styles part)?",
    ],
    bc_steps=[
        "Use the SAME sales quote as P01/P05/P13/P14; a second page (P15's long quote) shows the default header.",
        "Report Layouts > New layout for report 1304, upload this .docx (not template.dotx), then Print the "
        "quote with it.",
    ],
)
def p17(b: Build):
    info = b.info()
    for part in info["partDetails"]:
        if part["kind"] == "header" and part.get("role") in ("default", "first"):
            where = {"layoutPart": "header", "partName": part["name"]}
            title = b.edit("insert_label", label="/Header/DocumentTitle_Lbl", locationType="documentEnd",
                           format={"bold": True}, **where)["data"]["controlId"]
            b.edit("insert_field", field="/Header/DocumentNo", locationType="afterControl", controlId=title, **where)
            b.edit("insert_text", text=" ", locationType="afterControl", controlId=title, **where)
            b.note(f"{part['name']} ({part['role']} header): DocumentTitle_Lbl, a space and DocumentNo, bold")

    addr = b.edit("insert_table", rowCount=4, columns=widths_and_alignments([4600, 4600]),
                  locationType="documentEnd")["data"]
    for row in range(4):
        b.edit("insert_field", field=f"/Header/CompanyAddress{row + 1}", locationType="tableCell",
               tableIndex=addr["tableIndex"], row=row, col=0)
        b.edit("insert_field", field=f"/Header/CustomerAddress{row + 1}", locationType="tableCell",
               tableIndex=addr["tableIndex"], row=row, col=1)
    b.edit("insert_repeater_table", dataItem="/Header/Line", locationType="documentEnd", headerFromLabels=True,
           columns=repeater_columns(
               "/Header/Line", ["ItemNo_Line", "Description_Line", "Quantity_Line", "UnitPrice", "LineAmount_Line"],
               [1600, 4000, 1100, 1400, 1600], ["left", "left", "right", "right", "right"]))
    totals = b.edit("insert_table", rowCount=1, columns=widths_and_alignments([6200, 2400], ["right", "right"]),
                    locationType="documentEnd")["data"]
    b.edit("insert_field", field="/Header/Totals/TotalIncludingVATText", locationType="tableCell",
           tableIndex=totals["tableIndex"], row=0, col=0)
    b.edit("insert_field", field="/Header/Totals/TotalAmountIncludingVAT", locationType="tableCell",
           tableIndex=totals["tableIndex"], row=0, col=1, format={"bold": True})
    b.edit("set_cell_format", tableIndex=totals["tableIndex"], row=0, edges=["top"])
    b.note("body under create_layout's heading: a 2x4 address table, the /Header/Line repeater table and a "
           "1x2 totals block with a rule above - P01's recipe, here over the template")


# P18-P20 follow the round-3 re-run (2026-10-25): P18 checks why P14 printed no work description, and P19/P20
# carry #123 on after P09/P10 showed that neither hide-if-empty wrapper acts on a row outside a repeater whose
# bound record is missing - P19 with an empty value outside a repeater, P20 inside one.

def inline_work_description_field(xml: str) -> str:
    """The stock quote's work-description field, a BLOCK-level control around its paragraph, becomes the INLINE
    form P14 had: the paragraph outside, the field inside it (no tool writes that form any more)."""
    start = control_start(xml, xpath=WORK_DESCRIPTION_LINE_XPATH)
    end = element_end(xml, start, "sdt")
    parts = re.fullmatch(r"<w:sdt>(<w:sdtPr>.*?</w:sdtPr>)(<w:sdtEndPr>.*?</w:sdtEndPr>)?<w:sdtContent>"
                         r"(<w:p\b[^>]*>)((?:<w:pPr>.*?</w:pPr>|<w:pPr/>)?)(.*)</w:p></w:sdtContent></w:sdt>",
                         xml[start:end], re.S)
    if parts is None:
        raise ItemFailure("the stock work-description field is not a block-level control around one paragraph")
    sdt_pr, end_pr, p_open, p_pr, runs = parts.groups()
    inline = f"{p_open}{p_pr}<w:sdt>{sdt_pr}{end_pr or ''}<w:sdtContent>{runs}</w:sdtContent></w:sdt></w:p>"
    return xml[:start] + inline + xml[end:]


WORK_DESCRIPTION_LINE_XPATH = ("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:WorkDescriptionLines[1]/"
                               "ns0:WorkDescriptionLine[1]")
YOUR_REFERENCE_XPATH = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:YourReference[1]"


@item(
    id="p18", title="Work description held inline on the stock quote",
    report="1304 - Standard Sales Quote", base="StandardSalesQuote.docx",
    purpose="The first re-run printed no work description from P14, although the quote had one and Word's mock "
            "printed every line. P14's work-description repeater held its field INLINE in a paragraph; the stock "
            "quote's own holds a BLOCK-level field around the paragraph, as every nested block repeater in "
            "Microsoft's layouts does. This is the stock quote with only that field turned into the inline form - "
            "the control experiment. If the built-in layout prints the work description and this one does not, "
            "the inline form was the cause, and wrap_in_repeater's new block form (rebuilt P14) is the fix.",
    asks=[
        "Quote 1003 with its work description: does THIS layout print the work description (between the info "
        "grid and the lines table), as the built-in layout does?",
        "P14, rebuilt: does it print the work description below its totals now?",
    ],
    bc_steps=[
        "Keep the work description on quote 1003 (two or three lines).",
        "Report Layouts > New layout for report 1304, upload this .docx and the rebuilt P14.",
        "Print quote 1003 three times: with the BUILT-IN layout (the reference), with this layout, and with P14.",
    ],
)
def p18(b: Build):
    b.drop_broken_abn_control()
    parts = rewrite_parts(b.work, r"word/document\.xml", inline_work_description_field)
    b.note(f"hand edit in {', '.join(parts)}: the work-description repeater's field "
           "(/Header/WorkDescriptionLines/WorkDescriptionLine) turned from a block-level control around its paragraph "
           "into an inline control inside it - the form P14 printed nothing with; nothing else changed")


@item(
    id="p19", title="Hide Table Row If Empty around an empty value",
    report="1304 - Standard Sales Quote", base="StandardSalesQuote.docx",
    purpose="GitHub issue #123. P09 and P10 showed that BC 28 acts on neither hide-if-empty wrapper around a row "
            "outside a repeater whose bound record is MISSING (a quote with no lines). This asks about a value that "
            "EXISTS but is empty: the 'Hide Table Row If Empty' wrapper goes around the info grid's Your Reference "
            "value, empty on most Cronus quotes. In the stock layout that value is an inline field in a plain cell, "
            "so it is first turned into a cell-level control, the shape the captured wrapper encloses.",
    asks=[
        "Your Reference EMPTY: is the info grid's value row (the salesperson, payment terms and the other values) "
        "left out, while its caption row still prints?",
        "Your Reference filled in: is the value row printed normally?",
    ],
    bc_steps=[
        "Report Layouts > New layout for report 1304, upload this .docx.",
        "Print quote 1003 with it as it is (Your Reference empty), then fill in Your Reference on the quote and "
        "print it again.",
    ],
)
def p19(b: Build):
    b.drop_broken_abn_control()
    wrapper_id = wrap_in_hide_if_empty(b, *HIDE_ROW_WRAPPER, xpath=YOUR_REFERENCE_XPATH, lift=True)
    b.note(f"hand edit in word/document.xml: the info grid's Your Reference value (/Header/YourReference) turned into "
           f"a cell-level control and wrapped in a 'Hide Table Row If Empty' control (id {wrapper_id}), the captured "
           "shape")


# P20's three rows, per invoice line: the caption, then the dataset path of the one cell the wrapper encloses.
HIDE_INSIDE_REPEATER_ROWS = [
    ("P20 A - first assembly line (no record):", "/Header/Line/AssemblyLine/Description_AssemblyLine"),
    ("P20 B - line discount text (empty):", "/Header/Line/LineDiscountPercentText_Line"),
    ("P20 C - description (has a value):", "/Header/Line/Description_Line"),
]


@item(
    id="p20", title="Hide Table Row If Empty inside a repeater",
    report="1306 - Standard Sales Invoice", base="StandardSalesInvoiceVatSpec.docx",
    purpose="GitHub issue #123, inside a repeater. P09, P10 and P19 showed that BC 28 acts on neither "
            "hide-if-empty wrapper OUTSIDE a repeater, whether the bound record is missing or its value is empty. "
            "The one capture of the row wrapper sits in a repeater's row, so this asks there, on data every sandbox "
            "has: P08's invoice layout gets three rows under each invoice line (insert_subtotal_row), each with a "
            "caption and one cell wrapped in 'Hide Table Row If Empty' - first turned into a cell-level control, the "
            "captured shape. A binds the line's first ASSEMBLY line, a missing record on an ordinary invoice: the "
            "exact shape captured in Salesperson Commission. B binds the line's discount text, present but empty "
            "when the line has no discount. C binds its description, never empty: the control.",
    asks=[
        "Row A (P20 A - first assembly line): is it left out under every line, or printed with nothing after "
        "the caption?",
        "Row B (P20 B - line discount text): left out, or printed empty?",
        "Row C (P20 C - description): printed under every line, with the description? It must be: if it is "
        "missing too, the wrapper hides the row whatever the value.",
    ],
    bc_steps=[
        "The same posted invoice as P08 (any posted sales invoice with a few item lines and no line discounts "
        "works; it needs no assembly items).",
        "Report Layouts > New layout for report 1306, upload this .docx, then Print the posted invoice with it.",
    ],
)
def p20(b: Build):
    # P08's repair first: the stock capture's title and number bindings name a foreign namespace BC rejects.
    b.edit("refresh_xml_part",
           newSchemaSource=os.path.join(CORPUS_DIR, "StandardSalesInvoiceVatSpec.docx"))
    b.note("refresh_xml_part against the layout's own schema, as P08: the foreign-namespace bindings BC rejects "
           "on upload are re-pointed")

    info = b.info()
    lines = Build.repeater(info, "/Header/Line")
    table = Build.line_table(info, lines)
    grid = next(t for t in info["tables"]
                if t["tableIndex"] == table and t["part"] == "document.xml")["columnCount"]
    wrapped = []
    for caption, path in HIDE_INSIDE_REPEATER_ROWS:
        before = {c["sdtId"] for c in b.info()["controls"]}
        b.edit("insert_subtotal_row", parentControlId=lines["sdtId"],
               cells=[{"span": 4}, {"fields": [path], "span": 4}, {"span": grid - 8}])
        info = b.info()
        added = [c for c in info["controls"] if c["sdtId"] not in before]
        if len(added) != 1:
            raise ItemFailure(f"expected one new control bound to {path}, found {len(added)}")
        rows = next(t for t in info["tables"] if t["tableIndex"] == table and t["part"] == "document.xml")["rows"]
        row = next(r["rowIndex"] for r in rows
                   if any(added[0]["sdtId"] in (cell.get("innerControlIds") or []) for cell in r["cells"]))
        b.edit("set_cell_text", tableIndex=table, row=row, col=0, text=caption)
        wrapped.append((added[0]["sdtId"], caption.split(" - ")[0], path))
    b.note(f"insert_subtotal_row x3 on the /Header/Line repeater (+ set_cell_text captions): rows "
           f"{', '.join(label for _, label, _ in wrapped)} under each invoice line, binding "
           f"{', '.join(path for _, _, path in wrapped)}")

    for control_id, label, path in wrapped:
        wrapper_id = wrap_in_hide_if_empty(b, *HIDE_ROW_WRAPPER, control_id=control_id, lift=True)
        b.note(f"hand edit: {label}'s cell ({path}) turned into a cell-level control and wrapped in a 'Hide Table Row "
               f"If Empty' control (id {wrapper_id}), the shape captured in Salesperson Commission")


# ---------------------------------------------------------------- runner

# Business Central refuses to upload a layout whose settings link an external attached template:
# "UnsupportedExternalContentFields: ... ExternalRelationship:attachedTemplate" (a BC 28 sandbox, 2026-10-24).
# Rounds 1 and 2 uploaded the same stock bases, link included, so the check is newer than they are. Four stock
# captures carry one. Every tool that writes a layout removes it, so only an item no tool saves still carries
# one (P11, edited by hand); the pack unlinks it there before validation.
ATTACHED_TEMPLATE_REL = re.compile(r'<Relationship\b[^>]*/attachedTemplate"[^>]*/>')
ATTACHED_TEMPLATE_SETTING = re.compile(r"<w:attachedTemplate\b[^>]*/>")


def strip_attached_template(path: str) -> list[str]:
    """Unlinks the external attached template of the layout at path: its settings.xml.rels relationship and the
    w:attachedTemplate setting naming it. Returns the targets it unlinked, empty when there was none."""
    rels_part = "word/_rels/settings.xml.rels"
    with zipfile.ZipFile(path) as package:
        rels = package.read(rels_part).decode("utf-8") if rels_part in package.namelist() else ""
    targets = [re.search(r'Target="([^"]*)"', m.group(0)).group(1) for m in ATTACHED_TEMPLATE_REL.finditer(rels)]
    if targets:
        rewrite_parts(path, re.escape(rels_part), lambda xml: ATTACHED_TEMPLATE_REL.sub("", xml))
        rewrite_parts(path, r"word/settings\.xml", lambda xml: ATTACHED_TEMPLATE_SETTING.sub("", xml))
    return targets


def clear_generated(out_dir: str):
    """Deletes what an earlier build of the item wrote - its layout and template, the merged copies, the mock
    and its page images - so a rebuild that changes the report or the page count leaves nothing stale behind.
    A reviewer's own files (INSTRUCTIONS.md, steps 4-5) are kept."""
    for name in os.listdir(out_dir):
        path = os.path.join(out_dir, name)
        if name.endswith((".docx", ".dotx")) or name == "mock-preview.pdf":
            os.remove(path)
        elif name == "mock-pages" and os.path.isdir(path):
            shutil.rmtree(path)


def build_item(entry: dict, out_dir: str) -> dict:
    os.makedirs(out_dir, exist_ok=True)
    clear_generated(out_dir)
    name = entry["report"].split(" - ", 1)[-1].replace(" ", "")
    work = os.path.join(out_dir, f"{name}.docx")
    server = McpServer(env=entry["env"] or None)
    record: dict = {"id": entry["id"], "title": entry["title"], "report": entry["report"],
                    "purpose": entry["purpose"], "asks": entry["asks"], "bcSteps": entry["bc_steps"],
                    "env": entry["env"], "base": entry["base"] or f"(created from {entry['schema']}"
                    + (", over the template.dotx in this folder)" if entry["template"] else ")")}
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
            create = {"schemaSource": os.path.join(CORPUS_DIR, entry["schema"]), "outputPath": work,
                      "headingText": entry["report"].split(" - ", 1)[-1].upper()}
            if entry["template"]:
                create["templatePath"] = entry["template"](b, out_dir)
            b.tool("create_layout", create)
        entry["fn"](b)
        unlinked = strip_attached_template(work)
        if unlinked:
            b.note(f"hand edit: unlinked the external attached template {', '.join(unlinked)} that the layout "
                   "inherits from the stock layout it starts from - BC refuses to upload a layout linking one "
                   "(UnsupportedExternalContentFields), and no tool saved this item to remove it")

        validation = b.tool("validate_layout", {"layoutPath": work, "level": "full"})["data"]
        preview = b.tool("preview_layout", {"layoutPath": work, "sampleRows": PREVIEW_ROWS,
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
            render = server.tool("render_preview_pages", {"pdfPath": target, "maxPages": RENDER_PAGES},
                                 image_dir=os.path.join(out_dir, "mock-pages"))
            if not render.get("ok"):
                raise ItemFailure(f"render_preview_pages failed: {describe_error(render)}")
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

**Why:** only the real BC report engine can prove what this tool emits renders correctly. Rounds 1
(2026-08-01, five items) and 2 (2026-08-02, seven items, typography) passed and are recorded in
`docs/FIDELITY-CHECKLIST.md`; re-running P01-P07 whenever the tool surface changes is standing
verification debt (GitHub issue #14). 2.0 changed how several of the tools they use write (insert_column's
rows and widths, inserts taking the formatting around them), so this round re-runs them too.

**Round 3 adds four research probes, P08-P11.** They don't check a capability: each settles an open
issue that is blocked on what BC does - #102 (P08), #123 (P09, P10) and #11 (P11). For these, a
REJECTED upload or a missing row is an answer, not a failure: write down exactly what happened.
**P12-P14** are capability checks again, of what 2.0 writes: a statement repeated per customer by
`wrap_in_repeater` (P12), the formatting tools #15 added, including direct italic, which no stock layout
carries (P13), and the reworked table tools - tables in the header and footer, the three `borders`
looks, a rule under every line - with the work description repeated per line (P14).
**P15-P17** print what no round has printed from the tools at all: totals rows inside a lines table and
page numbers (P15), per-group subtotals (P16), and a layout created over a branded `.dotx` template (P17).

**Round 3's re-run (2.0.0-rc.2)** follows the first round-3 prints (2026-10-24). P04 and P06 are corrected
(P04's column order and typography, P06's vendor address), P08 is a regression check of the #102 fix, and
P09/P10 moved to the stock quote, where a quote with no lines gives the hide-if-empty wrappers an empty case.
P11, P14 and P15 are unchanged, but their round-3 prints missed a case each: a one-page order (P11), a quote
with a work description (P14) and a quote long enough for two pages (P15).

**P18-P20 follow that re-run.** P14 printed no work description although the quote had one, so it is rebuilt
with `wrap_in_repeater`'s new block-level form, and P18 prints the old inline form on the stock quote to confirm
the cause. P09/P10 showed that neither hide-if-empty wrapper acts on a row outside a repeater whose bound record
is missing; P19 tries a value that exists but is empty, and P20 the captured shape inside a repeater (#123).

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
gets filled in per item, and the real findings go to `docs/FIDELITY-CHECKLIST.md` and the issue tracker.
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


# What a reviewer adds to an item folder (INSTRUCTIONS.md, steps 4-5, and the dataset tip).
REVIEWER_FILES = re.compile(r".*-bc(-stock)?\.pdf|.*-dataset\.xml", re.IGNORECASE)


def remove_stale_item_dirs(item_id: str, slug: str):
    """Deletes the folders an item built under an earlier title, so the pack holds one folder per item. A
    folder holding a reviewer's BC PDF or dataset is kept, with a warning."""
    if not os.path.isdir(OUT_ROOT):
        return
    for name in os.listdir(OUT_ROOT):
        path = os.path.join(OUT_ROOT, name)
        if name == slug or not name.startswith(f"{item_id}-") or not os.path.isdir(path):
            continue
        kept = [f for f in os.listdir(path) if REVIEWER_FILES.fullmatch(f)]
        if kept:
            print(f"    kept the stale folder {name}: it holds {', '.join(kept)}")
        else:
            shutil.rmtree(path)
            print(f"    removed the stale folder {name}")


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
        remove_stale_item_dirs(e["id"], slug)
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
