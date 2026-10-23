using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #80: <see cref="TableGridNavigator.Rows"/> yielded one row per direct child of the
/// <c>w:tbl</c>, resolving a row-level control to the FIRST <c>w:tr</c> inside it. Every other row of a
/// repeater item — nested detail repeaters, subtotal and spacer rows, the extra rows of a multi-row item — was
/// invisible, to the column tools, the width re-slicing, the vMerge refusal, <c>get_layout_info</c>,
/// <c>tableCell</c> addressing and <see cref="TableGridConsistencyGuard"/> alike. A column edit therefore left
/// those rows on the old grid while the guard built to reject exactly that saw nothing.
/// </summary>
/// <remarks>
/// <para>
/// Multi-row items are the corpus norm (InventoryOrderDetails, StandardStatement,
/// SalesInvoiceForSubscriptionBilling, SalespersonCommission, SubcontractorDispatchList), and the tools make
/// the shape themselves (<c>insert_repeater_row</c>, <c>insert_subtotal_row</c>). Every assertion about the
/// result walks the rows with <see cref="IndependentGridWalk"/>, not the navigator, so these tests cannot
/// share the blind spot they are about.
/// </para>
/// <para>
/// InventoryOrderDetails, table 0, for reference: rows 0–2 are plain; the <c>/Item</c> repeater
/// (<see cref="ItemRepeaterId"/>) holds rows 3 (spacer), 4 (item fields), 5 (the nested
/// <c>/Item/Sales_Line</c> repeater's row) and 6 (subtotals); rows 7–8 are plain totals rows.
/// </para>
/// </remarks>
public class NestedRepeaterRowTests
{
    private const int ItemRepeaterId = 1709525787;       // #Nav: /Item
    private const int ItemRepeaterItemId = -1844079213;  // its repeatingSectionItem
    private const int SalesLineRepeaterId = 1689722507;  // #Nav: /Item/Sales_Line
    private const int SalesLineItemId = 1248918833;      // its repeatingSectionItem
    private const int SalesLineAmountId = -2047056434;   // /Item/Sales_Line/OutstandingAmt1_SalesLine, grid column 9

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-nestedrows-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static Table BodyTable(WordprocessingDocument doc, int tableIndex) =>
        TableGridNavigator.Tables(doc.MainDocumentPart!.Document!.Body!)[tableIndex];

    private static void AssertEveryRowCoversTheGrid(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        Assert.Empty(IndependentGridWalk.RowsOffGrid(doc));
    }

    // ---- the navigator itself ----

    [Theory]
    [InlineData(Corpus.InventoryOrderDetails, 0)]
    [InlineData(Corpus.StandardStatement, 3)]
    [InlineData(Corpus.StandardStatement, 4)]
    [InlineData(Corpus.SalesInvoice, 3)]
    [InlineData(Corpus.SalesInvoice, 5)]
    [InlineData(Corpus.SalespersonCommission, 0)]
    [InlineData(Corpus.SubcontractorDispatchList, 0)]
    public void Rows_lists_every_w_tr_of_the_table_in_document_order(string corpusFile, int tableIndex)
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(corpusFile), false);
        var table = BodyTable(doc, tableIndex);

        var rows = TableGridNavigator.Rows(table);

        Assert.Equal(IndependentGridWalk.Rows(table), rows.Select(r => r.InnerRow!).ToList());
    }

    [Fact]
    public void Rows_records_the_whole_wrapper_chain_of_a_row_nested_two_repeaters_deep()
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.InventoryOrderDetails), false);
        var table = BodyTable(doc, 0);

        var row = TableGridNavigator.Rows(table)[5];

        Assert.True(row.IsControlRow);
        Assert.Equal(ItemRepeaterId, row.ControlId); // the outermost wrapper, as for the item's first row
        Assert.Equal(
            new int?[] { ItemRepeaterId, ItemRepeaterItemId, SalesLineRepeaterId, SalesLineItemId },
            row.Wrappers.Select(w => SdtInspector.ReadControlId(w)).ToArray());
        Assert.Same(row.Wrappers[0], row.RowChild);
        Assert.Same(table, row.RowChild.Parent);
    }

    // ---- the issue's three reproductions ----

    [Fact]
    public void Remove_column_updates_every_row_of_a_multi_row_repeater_item()
    {
        // e2e s16's call. It reported "updating 6 row(s)" and left the three rows nested in the /Item item —
        // the real /Item/Sales_Line data row among them — on ten cells, the removed amount printing past the
        // table's right edge.
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        try
        {
            var response = TableTools.RemoveColumn(path, tableIndex: 0, column: 9);
            Assert.True(response.Ok, response.Error?.Message);
            Assert.Contains("updating 9 row(s)", Assert.IsType<TableEditResultDto>(response.Data).Summary);

            AssertEveryRowCoversTheGrid(path);
            using var doc = WordprocessingDocument.Open(path, false);
            Assert.DoesNotContain(
                doc.MainDocumentPart!.Document!.Body!.Descendants<SdtElement>(),
                s => SdtInspector.ReadControlId(s) == SalesLineAmountId);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Remove_column_updates_the_nested_detail_row_and_the_trailing_item_row_of_a_statement()
    {
        // e2e s10's call: "updating 2 row(s)", with the nested DtldCustLedgEntries row and the item's
        // trailing row still covering 8 of 7 columns.
        var path = CopyOfCorpus(Corpus.StandardStatement);
        try
        {
            var response = TableTools.RemoveColumn(path, tableIndex: 3, column: 7);
            Assert.True(response.Ok, response.Error?.Message);
            Assert.Contains("updating 4 row(s)", Assert.IsType<TableEditResultDto>(response.Data).Summary);

            AssertEveryRowCoversTheGrid(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Insert_column_reaches_the_detail_spacer_and_subtotal_rows_the_tools_themselves_built()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var table = Assert.IsType<EditResultDto>(TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line,Description_Line,Quantity_Line,LineAmount_Line", "documentEnd").Data);
            var tableIndex = table.TableIndex!.Value;
            Assert.True(TableTools.InsertRepeaterRow(
                path, table.ControlId, "/Header/Line/AssemblyLine", cells: "-,2:Description_AssemblyLine,-").Ok);
            Assert.True(TableTools.InsertSubtotalRow(path, table.ControlId, cells: "4:-").Ok);
            Assert.True(TableTools.InsertSubtotalRow(
                path, table.ControlId,
                cells: "2:-,/Header/Line/AmountExcludingVAT_Line_Lbl,/Header/Line/AmountExcludingVAT_Line").Ok);

            var response = TableTools.InsertColumn(path, tableIndex, mode: "field", dataPath: "/Header/Line/UnitPrice");
            Assert.True(response.Ok, response.Error?.Message);
            Assert.Contains("adding a cell to 5 row(s)", Assert.IsType<TableEditResultDto>(response.Data).Summary);

            AssertEveryRowCoversTheGrid(path);

            // The field went into the line row only - not into the detail row, the spacer or the subtotal
            // row, which are filler for a /Header/Line column.
            using var doc = WordprocessingDocument.Open(path, false);
            var rows = IndependentGridWalk.Rows(BodyTable(doc, tableIndex));
            Assert.Equal(5, rows.Count);
            var bound = rows
                .Select((r, i) => (Index: i, Count: r.Descendants<SdtElement>().Count(s => IsUnitPrice(SdtInspector.ReadXPath(s)))))
                .Where(x => x.Count > 0)
                .ToList();
            Assert.Equal((1, 1), Assert.Single(bound));
        }
        finally
        {
            File.Delete(path);
        }

        static bool IsUnitPrice(string? xpath) =>
            xpath is not null
            && (xpath.EndsWith("ns0:UnitPrice[1]", StringComparison.Ordinal) || xpath.EndsWith("ns0:UnitPrice", StringComparison.Ordinal));
    }

    [Fact]
    public void Insert_column_binds_an_item_field_into_the_row_holding_its_siblings_not_the_items_leading_spacer()
    {
        // The /Item item opens with an empty spacer row (3); its own fields are on row 4. The field goes to
        // ONE row per item - the one its sibling fields sit in - and every other row of the item gets filler.
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        try
        {
            var response = TableTools.InsertColumn(
                path, tableIndex: 0, mode: "field", dataPath: "/Item/InventoryPostingGroup_Item");
            Assert.True(response.Ok, response.Error?.Message);

            AssertEveryRowCoversTheGrid(path);
            using var doc = WordprocessingDocument.Open(path, false);
            Assert.Equal(new[] { 4 }, RowsBinding(BodyTable(doc, 0), "ns0:InventoryPostingGroup_Item"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Insert_column_still_takes_a_parent_field_when_the_only_other_repeater_is_a_nested_detail_one()
    {
        // A /Header field is placed only in a single-repeater table. The detail repeater nested in the
        // /Header/Line item is now visible, but it is part of that item, not a second repeater to choose from.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var table = Assert.IsType<EditResultDto>(TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line,Description_Line", "documentEnd").Data);
            Assert.True(TableTools.InsertRepeaterRow(
                path, table.ControlId, "/Header/Line/AssemblyLine", cells: "-,Description_AssemblyLine").Ok);

            var response = TableTools.InsertColumn(
                path, table.TableIndex!.Value, mode: "field", dataPath: "/Header/BillToContactEmail");
            Assert.True(response.Ok, response.Error?.Message);

            AssertEveryRowCoversTheGrid(path);
            using var doc = WordprocessingDocument.Open(path, false);
            Assert.Equal(new[] { 1 }, RowsBinding(BodyTable(doc, table.TableIndex!.Value), "ns0:BillToContactEmail"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>Indexes (in <see cref="IndependentGridWalk.Rows"/> order) of the rows holding a control whose XPath's last step is <paramref name="leafStep"/>.</summary>
    private static int[] RowsBinding(Table table, string leafStep) =>
        IndependentGridWalk.Rows(table)
            .Select((row, index) => (row, index))
            .Where(x => x.row.Descendants<SdtElement>().Any(s =>
                SdtInspector.ReadXPath(s) is { } xpath
                && (xpath.EndsWith(leafStep + "[1]", StringComparison.Ordinal) || xpath.EndsWith(leafStep, StringComparison.Ordinal))))
            .Select(x => x.index)
            .ToArray();

    // ---- the other operations built on the same row walk ----

    [Fact]
    public void Set_column_widths_re_slices_the_cells_of_every_nested_row()
    {
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        try
        {
            var widths = Enumerable.Range(0, 10).Select(i => 1000 + (i * 10)).ToArray();
            var response = TableTools.SetColumnWidths(path, tableIndex: 0, widths);
            Assert.True(response.Ok, response.Error?.Message);
            Assert.Contains("resized 9 row(s)", Assert.IsType<TableEditResultDto>(response.Data).Summary);

            using var doc = WordprocessingDocument.Open(path, false);
            foreach (var row in IndependentGridWalk.Rows(BodyTable(doc, 0)))
            {
                var cursor = 0;
                foreach (var cell in IndependentGridWalk.Cells(row))
                {
                    var span = IndependentGridWalk.Span(cell);
                    Assert.Equal(
                        widths.Skip(cursor).Take(span).Sum().ToString(System.Globalization.CultureInfo.InvariantCulture),
                        cell.TableCellProperties!.TableCellWidth!.Width!.Value);
                    cursor += span;
                }
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Set_column_widths_reaches_the_wrapped_spanned_cell_deep_in_a_five_level_chain()
    {
        // SubcontractorDispatchList's "Hide Table If Empty" cell (issue #85) sits in a row nested four
        // repeaters deep - the two fixes together are what let a width edit reach it.
        var path = CopyOfCorpus(Corpus.SubcontractorDispatchList);
        try
        {
            var widths = Enumerable.Range(0, 10).Select(i => 900 + (i * 10)).ToArray();
            Assert.True(TableTools.SetColumnWidths(path, tableIndex: 0, widths).Ok);

            using var doc = WordprocessingDocument.Open(path, false);
            var wrapper = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtCell>()
                .Single(s => SdtInspector.ReadAlias(s) == "Hide Table If Empty");
            var cell = wrapper.Descendants<TableCell>().First();
            Assert.Equal(widths.Skip(2).Take(4).Sum().ToString(System.Globalization.CultureInfo.InvariantCulture),
                cell.TableCellProperties!.TableCellWidth!.Width!.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void The_vMerge_refusal_sees_a_vertical_merge_in_a_nested_row()
    {
        var path = SyntheticLayout.Create(MultiRowItemTable(secondItemRowCell:
            "<w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p/></w:tc>"));
        try
        {
            var response = TableTools.InsertColumn(path, tableIndex: 0, mode: "plainText");

            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.Contains("w:vMerge", response.Error.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void The_grid_guard_sees_an_off_grid_row_nested_in_a_repeater_item()
    {
        // The item's SECOND row covers one of two grid columns: the guard used to report nothing.
        var path = SyntheticLayout.Create(MultiRowItemTable(secondItemRowCell: string.Empty));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var violation = Assert.Single(TableGridConsistencyGuard.Find(doc));

            Assert.Equal(TableGridViolationKind.CoverageMismatch, violation.Kind);
            Assert.Equal(2, violation.RowIndex);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- addressing a nested row by index ----

    [Fact]
    public void Get_layout_info_lists_the_nested_rows_and_gives_their_controls_table_coordinates()
    {
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        try
        {
            var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            var table = info.Tables.Single(t => t.Part == "document.xml" && t.TableIndex == 0);

            Assert.Equal(9, table.RowCount);
            Assert.Equal(Enumerable.Range(0, 9), table.Rows.Select(r => r.RowIndex));
            Assert.Equal(SalesLineAmountId, table.Rows[5].Cells[9].ControlId);

            var amount = Assert.Single(info.Controls, c => c.SdtId == SalesLineAmountId);
            Assert.Equal((0, 5, 9), (amount.TableIndex, amount.RowIndex, amount.ColIndex));

            // A row-level control's coordinate is the FIRST row it wraps.
            Assert.Equal(3, Assert.Single(info.Controls, c => c.SdtId == ItemRepeaterId).RowIndex);
            Assert.Equal(5, Assert.Single(info.Controls, c => c.SdtId == SalesLineRepeaterId).RowIndex);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Get_layout_info_tags_each_row_with_the_innermost_repeater_it_belongs_to()
    {
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        try
        {
            var table = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data).Tables
                .Single(t => t.Part == "document.xml" && t.TableIndex == 0);

            Assert.Equal(
                new int?[] { null, null, null, ItemRepeaterId, ItemRepeaterId, SalesLineRepeaterId, ItemRepeaterId, null, null },
                table.Rows.Select(r => r.RepeaterControlId).ToArray());

            // ControlId stays the OUTERMOST control, the one whose removal takes every row of the item.
            Assert.All(table.Rows.Skip(3).Take(4), r => Assert.Equal(ItemRepeaterId, r.ControlId));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Cell_tools_address_a_cell_in_a_row_after_the_first_of_a_repeater_item()
    {
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        try
        {
            // Row 6 is the /Item item's subtotal row; its cell 2 is a plain, empty cell.
            Assert.True(EditTools.SetCellText(path, tableIndex: 0, row: 6, col: 2, text: "Subtotal").Ok);
            Assert.True(TableTools.SetCellFormat(path, tableIndex: 0, row: 6, edges: ["top"]).Ok);
            Assert.True(TableTools.MergeCells(path, tableIndex: 0, row: 6, fromColumn: 2, toColumn: 3).Ok);

            using var doc = WordprocessingDocument.Open(path, false);
            var subtotalRow = IndependentGridWalk.Rows(BodyTable(doc, 0))[6];
            var cells = IndependentGridWalk.Cells(subtotalRow);
            Assert.Equal("Subtotal", cells[2].InnerText);
            Assert.Equal(2, IndependentGridWalk.Span(cells[2]));
            Assert.All(cells, c =>
                Assert.Equal(BorderValues.Single, c.TableCellProperties!.TableCellBorders!.TopBorder!.Val!.Value));
            Assert.Empty(IndependentGridWalk.RowsOffGrid(doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(4)] // between the item's spacer row (3) and its field row (4)
    [InlineData(6)] // between the nested Sales_Line row (5) and the subtotal row (6)
    public void Insert_table_row_refuses_a_position_inside_a_repeater(int atRow)
    {
        // A static row there would have to go inside the repeating item (and repeat per item) or be moved out
        // of place - neither is what "insert at this row" asked for.
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        try
        {
            var before = File.ReadAllBytes(path);

            var response = TableTools.InsertTableRow(path, tableIndex: 0, cells: "10:-", atRow: atRow);

            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.Contains("inside", response.Error.Message);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(3, 3)] // just before the /Item repeater
    [InlineData(7, 7)] // just after it
    public void Insert_table_row_at_a_repeater_boundary_lands_outside_the_repeater(int atRow, int expectedIndex)
    {
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        try
        {
            var response = TableTools.InsertTableRow(path, tableIndex: 0, cells: "10:-", atRow: atRow);
            Assert.True(response.Ok, response.Error?.Message);

            using var doc = WordprocessingDocument.Open(path, false);
            var table = BodyTable(doc, 0);
            var rows = IndependentGridWalk.Rows(table);
            Assert.Equal(10, rows.Count);
            Assert.Same(table, rows[expectedIndex].Parent); // a direct table row, so it renders once
            Assert.Empty(IndependentGridWalk.RowsOffGrid(doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// A two-column table: a plain header row, then a repeater whose item holds TWO rows. The first item row is
    /// well-formed; the second's second cell is <paramref name="secondItemRowCell"/> (empty = the row covers
    /// one grid column of two).
    /// </summary>
    private static string MultiRowItemTable(string secondItemRowCell)
    {
        const string Cell = "<w:tc><w:tcPr/><w:p/></w:tc>";
        return "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:tr>" + Cell + Cell + "</w:tr>"
            + "<w:sdt><w:sdtPr><w:id w:val=\"701\"/><w15:repeatingSection/></w:sdtPr><w:sdtContent>"
            + "<w:sdt><w:sdtPr><w:id w:val=\"702\"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>"
            + "<w:tr>" + Cell + Cell + "</w:tr>"
            + "<w:tr>" + Cell + secondItemRowCell + "</w:tr>"
            + "</w:sdtContent></w:sdt>"
            + "</w:sdtContent></w:sdt>"
            + "</w:tbl>";
    }
}
