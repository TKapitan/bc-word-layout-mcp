"""Author a complete Sales Order Confirmation layout from scratch via MCP tools only.

The reference script for from-scratch document authoring, using the full 2026-07-31 toolset:
create_layout -> title band with a company-logo picture placeholder (insert_picture) -> plain
address/info tables (insert_table) filled with bound fields/labels (bold 8pt captions) -> a
BC-native-look lines repeater (no grid, header rule) with right-aligned numeric columns -> a
right-anchored totals block with real rules (set_cell_borders) -> content in REAL header/footer
parts (auto-scaffolded) -> an interior-position insert_column -> a NESTED AssemblyLine repeater
expanding once per line. Requires the Release build (dotnet build -c Release). Output:
preview-output/from-scratch-order-conf/SalesOrderConfirmation.docx (preview it with preview_layout).

Usage: python tools/e2e/author_order_confirmation.py [through-step]   (3..10; default: everything)
"""
import json, subprocess, sys, os

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..")).replace("\\", "/")
WORK = f"{ROOT}/preview-output/from-scratch-order-conf/SalesOrderConfirmation.docx"
# Schema source only - create_layout copies this layout's BC dataset part and builds a blank body from
# it. Standard_Sales_Invoice/1306 is used because it carries every path this script authors, including
# the nested /Header/Line/AssemblyLine detail item that step 10 needs.
CORPUS = f"{ROOT}/tests/corpus/SalesInvoiceForSubscriptionBilling.docx"
THROUGH = int(sys.argv[1]) if len(sys.argv) > 1 else 99

def tool(name, args):
    out = subprocess.run([sys.executable, os.path.join(HERE, "call.py"), name, json.dumps(args)],
                         capture_output=True, text=True, timeout=300)
    r = json.loads(out.stdout)
    if not r.get("ok"):
        print(f"FAIL {name}: {json.dumps(r.get('error'))[:300]}"); sys.exit(1)
    return r["data"]

def edit(name, **kw):
    return tool(name, {"layoutPath": WORK, **kw})

# ---- steps 1+2: create blank (the title goes into the logo band below, not the heading) ----
tool("create_layout", {"schemaSource": CORPUS, "outputPath": WORK, "headingText": ""})
print("created (blank body)")
if THROUGH < 3: sys.exit(0)

# ---- step 3: title band - document title left, company-logo picture placeholder right ----
t = edit("insert_table", rows=1, columns=2, columnWidths="7206,3000",
         locationType="documentEnd", columnAlignments="left,right")
edit("set_cell_text", tableIndex=t["tableIndex"], row=0, col=0,
     text="Sales Order Confirmation", bold=True, fontSizePoints=18)
edit("insert_picture", picture="/Header/CompanyPicture", locationType="tableCell",
     tableIndex=t["tableIndex"], row=0, col=1, widthMm=30, heightMm=20)
print("title band + logo")
if THROUGH < 4: sys.exit(0)

# ---- step 4: address block (customer left / company right) ----
t = edit("insert_table", rows=8, columns=2, locationType="documentEnd", columnAlignments="left,right")
for i in range(8):
    edit("insert_field", field=f"/Header/CustomerAddress{i+1}", locationType="tableCell",
         tableIndex=t["tableIndex"], row=i, col=0)
    edit("insert_field", field=f"/Header/CompanyAddress{i+1}", locationType="tableCell",
         tableIndex=t["tableIndex"], row=i, col=1)
print("addresses")
if THROUGH < 5: sys.exit(0)

# ---- step 5: header info grid (two label/field row pairs, bold 8pt captions) ----
t = edit("insert_table", rows=4, columns=4, locationType="documentEnd")
grid = [
    (0, ["/Header/YourReference_Lbl", "/Header/SelltoCustomerNo_Lbl", "/Header/SalesPerson_Lbl",
         "/Header/PaymentTermsDescription_Lbl"]),
    (1, ["/Header/YourReference", "/Header/SelltoCustomerNo", "/Header/SalesPersonName",
         "/Header/PaymentTermsDescription"]),
    (2, ["/Header/DocumentNo_Lbl", "/Header/DocumentDate_Lbl", "/Header/ShipmentDate_Lbl",
         "/Header/ShipmentMethodDescription_Lbl"]),
    (3, ["/Header/DocumentNo", "/Header/DocumentDate", "/Header/ShipmentDate",
         "/Header/ShipmentMethodDescription"]),
]
for row, paths in grid:
    for col, path in enumerate(paths):
        if path.endswith("_Lbl"):
            edit("insert_label", label=path, locationType="tableCell",
                 tableIndex=t["tableIndex"], row=row, col=col, bold=True, fontSizePoints=8)
        else:
            edit("insert_field", field=path, locationType="tableCell",
                 tableIndex=t["tableIndex"], row=row, col=col)
print("info grid")
if THROUGH < 6: sys.exit(0)

# ---- step 6: lines repeater - BC-native look (no grid, rule under the header row) ----
lines = edit("insert_repeater_table", dataItem="/Header/Line",
             columns="ItemNo_Line,Description_Line,Quantity_Line,UnitOfMeasure,UnitPrice,VATPct_Line,LineAmount_Line",
             locationType="documentEnd",
             columnWidths="1100,3100,900,900,1400,1000,1806",
             columnAlignments="left,left,right,left,right,right,right")
LINES_TABLE, DATA_ROW, LINES_REPEATER = lines["tableIndex"], lines["dataRowIndex"], lines["controlId"]
print(f"lines (bc look) tableIndex={LINES_TABLE} dataRowIndex={DATA_ROW} repeaterId={LINES_REPEATER}")
if THROUGH < 7: sys.exit(0)

# ---- step 7: totals block, right-anchored, with the corpus rules ----
t = edit("insert_table", rows=3, columns=3, columnWidths="5400,3000,1806",
         locationType="documentEnd", columnAlignments="left,left,right")
TOTALS_TABLE = t["tableIndex"]
totals = [
    (0, "/Header/Totals/TotalExcludingVATText", "/Header/Totals/TotalNetAmount", False),
    (1, "/Header/Totals/TotalVATAmountText", "/Header/Totals/TotalVATAmount", False),
    (2, "/Header/Totals/TotalIncludingVATText", "/Header/Totals/TotalAmountIncludingVAT", True),
]
for row, caption, amount, emphasize in totals:
    edit("insert_field", field=caption, locationType="tableCell",
         tableIndex=TOTALS_TABLE, row=row, col=1, bold=emphasize)
    edit("insert_field", field=amount, locationType="tableCell",
         tableIndex=TOTALS_TABLE, row=row, col=2, bold=emphasize)
# The rules a real BC totals block carries: one above the block, one above and below the grand total.
for col in (1, 2):
    edit("set_cell_borders", tableIndex=TOTALS_TABLE, row=0, col=col, edges="top")
    edit("set_cell_borders", tableIndex=TOTALS_TABLE, row=2, col=col, edges="top,bottom")
print("totals + rules")
if THROUGH < 8: sys.exit(0)

# ---- step 8: REAL page header/footer (auto-scaffolded on a from-scratch layout) ----
# Header: document no, so page 2+ still identifies the document. Stacked caption-over-value
# paragraphs (chaining two inline controls would concatenate them with no separator - there is no
# static-inline-text tool yet; see GitHub issue #5).
edit("insert_label", label="/Header/DocumentNo_Lbl", locationType="documentEnd",
     layoutPart="header", bold=True, fontSizePoints=8)
edit("insert_field", field="/Header/DocumentNo", locationType="documentEnd",
     layoutPart="header", fontSizePoints=8)
# Footer: the legal statement + VAT registration, repeating on every page.
edit("insert_field", field="/Header/CompanyLegalStatement", locationType="documentEnd",
     layoutPart="footer", fontSizePoints=8)
edit("insert_field", field="/Header/CompanyVATRegNo", locationType="documentEnd",
     layoutPart="footer", fontSizePoints=8)
# Body contact strip (a footer cannot hold an insert_table grid yet - see GitHub issue #10).
t = edit("insert_table", rows=2, columns=4, locationType="documentEnd")
contact = [
    ("/Header/CompanyPhoneNo_Lbl", "/Header/CompanyPhoneNo"),
    ("/Header/EMail_Header_Lbl", "/Header/CompanyEMail"),
    ("/Header/HomePage_Header_Lbl", "/Header/CompanyHomePage"),
    ("/Header/CompanyVATRegNo_Lbl", "/Header/CompanyVATRegNo"),
]
for col, (label, field) in enumerate(contact):
    edit("insert_label", label=label, locationType="tableCell",
         tableIndex=t["tableIndex"], row=0, col=col, bold=True, fontSizePoints=8)
    edit("insert_field", field=field, locationType="tableCell",
         tableIndex=t["tableIndex"], row=1, col=col, fontSizePoints=8)
print("page header/footer + contact strip")
if THROUGH < 9: sys.exit(0)

# ---- step 9: interior insert_column - Item Reference between Description and Quantity ----
edit("insert_column", tableIndex=LINES_TABLE, mode="field",
     dataPath="/Header/Line/ItemReferenceNo_Line",
     headerLabelPath="/Header/Line/ItemReferenceNo_Line_Lbl",
     atColumn=2, width=1200)
edit("set_column_widths", tableIndex=LINES_TABLE,
     widths="1100,2400,1200,800,800,1300,900,1706")  # re-fit to the 10206 content width
print("interior Item Reference column")
if THROUGH < 10: sys.exit(0)

# ---- step 10: NESTED detail rows - assembly components as lines UNDER each order line ----
# The standard BC shape: a detail ROW inside the outer repeater's item, aligned to the same grid
# (indented one column, description wide, quantity + unit right where the line's own numbers sit).
edit("insert_repeater_row", parentControlId=LINES_REPEATER,
     dataItem="/Header/Line/AssemblyLine",
     cells="-,3:Description_AssemblyLine,1:Quantity_AssemblyLine,2:UnitOfMeasure_AssemblyLine,1:-",
     alignments="-,left,right,left,-")
print("nested AssemblyLine detail rows")
print("AUTHORING COMPLETE")
