using BcWordLayout.Domain;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Direct coverage of <see cref="TableGridNavigator"/> — the single source of truth
/// <see cref="TableStructureReader"/> and <see cref="LocationResolver"/> now route through instead of
/// keeping their own row/cell-walking copies. With one implementation,
/// "parity" is a structural non-issue rather than a test to write; what actually needs direct coverage is
/// the navigator itself handling the tricky OOXML shapes real BC layouts produce: sdt-wrapped rows, a
/// multi-level row-control wrapper chain (a repeater's <c>repeatingSection</c> wrapping a
/// <c>repeatingSectionItem</c>), sdt-wrapped cells, <c>gridSpan</c>, <c>gridBefore</c>/<c>gridAfter</c>, and
/// a table nested inside an sdt block.
/// </summary>
public class TableGridNavigatorTests
{
    private static Table FirstTable(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        return doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().First();
    }

    // ---- Rows(): sees through a row-level sdt wrapper ----

    [Fact]
    public void Rows_returns_bare_rows_unwrapped_with_no_control_wrapper()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 1));
        try
        {
            var table = FirstTable(path);
            var rows = TableGridNavigator.Rows(table);

            Assert.Equal(2, rows.Count);
            Assert.All(rows, r =>
            {
                Assert.False(r.IsControlRow);
                Assert.Null(r.ControlId);
                Assert.Empty(r.Wrappers);
                Assert.Same(r.RowChild, r.InnerRow);
            });
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Rows_sees_through_a_row_level_sdt_wrapper_and_resolves_its_inner_w_tr()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithRowLevelControl(id: 99, text: "row-ctl"));
        try
        {
            var table = FirstTable(path);
            var rows = TableGridNavigator.Rows(table);

            var row = Assert.Single(rows);
            Assert.True(row.IsControlRow);
            Assert.Equal(99, row.ControlId);
            Assert.IsType<SdtRow>(row.RowChild);
            Assert.NotNull(row.InnerRow);
            Assert.Single(row.Wrappers);
            Assert.Same(row.RowChild, row.Wrappers[0]);
            Assert.Contains(row.InnerRow!.Descendants<Text>(), t => t.Text == "row-ctl");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Rows_records_every_level_of_a_nested_row_control_chain_via_ResolveRowWrapperChain()
    {
        // Mirrors a real BC repeater: an outer w15:repeatingSection SdtRow wraps a w15:repeatingSectionItem
        // SdtRow, which in turn wraps the data w:tr. TableStructureReader needs BOTH wrapper ids attributed
        // to the same row coordinate (a control-inventory lookup on either id must resolve).
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:sdt><w:sdtPr><w:id w:val=\"501\"/><w15:repeatingSection/></w:sdtPr><w:sdtContent>"
            + "<w:sdt><w:sdtPr><w:id w:val=\"502\"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>"
            + "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>row-item</w:t></w:r></w:p></w:tc></w:tr>"
            + "</w:sdtContent></w:sdt>"
            + "</w:sdtContent></w:sdt>"
            + "</w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var table = FirstTable(path);
            var rows = TableGridNavigator.Rows(table);

            var row = Assert.Single(rows);
            Assert.True(row.IsControlRow);
            Assert.Equal(501, row.ControlId); // the OUTERMOST wrapper's id
            Assert.Equal(2, row.Wrappers.Count);
            Assert.Equal(501, ReadId(row.Wrappers[0]));
            Assert.Equal(502, ReadId(row.Wrappers[1]));
            Assert.Contains(row.InnerRow!.Descendants<Text>(), t => t.Text == "row-item");

            // ResolveRowWrapperChain independently agrees with what Rows() already computed.
            var chain = TableGridNavigator.ResolveRowWrapperChain((SdtRow)row.RowChild);
            Assert.Equal(row.Wrappers, chain.Wrappers);
            Assert.Same(row.InnerRow, chain.InnerRow);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Rows_reports_null_InnerRow_when_a_row_level_wrapper_holds_no_w_tr()
    {
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:sdt><w:sdtPr><w:id w:val=\"777\"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>not-a-row</w:t></w:r></w:p></w:sdtContent></w:sdt>"
            + "</w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var table = FirstTable(path);
            var row = Assert.Single(TableGridNavigator.Rows(table));

            Assert.True(row.IsControlRow);
            Assert.Null(row.InnerRow);
            Assert.Single(row.Wrappers);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- Cells(): sees through a cell-level sdt wrapper, reports gridSpan ----

    [Fact]
    public void Cells_returns_bare_cells_unwrapped_with_default_gridSpan_one()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 2));
        try
        {
            var table = FirstTable(path);
            var innerRow = TableGridNavigator.Rows(table).Single().InnerRow!;
            var cells = TableGridNavigator.Cells(innerRow);

            Assert.Equal(2, cells.Count);
            Assert.All(cells, c =>
            {
                Assert.Same(c.CellChild, c.InnerCell);
                Assert.Equal(1, c.GridSpan);
            });
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Cells_sees_through_a_cell_level_sdt_wrapper_and_resolves_its_inner_w_tc()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithCellLevelControl(id: 42, text: "cell-ctl"));
        try
        {
            var table = FirstTable(path);
            var innerRow = TableGridNavigator.Rows(table).Single().InnerRow!;
            var cells = TableGridNavigator.Cells(innerRow);

            Assert.Equal(2, cells.Count);
            Assert.IsType<SdtCell>(cells[0].CellChild);
            Assert.NotNull(cells[0].InnerCell);
            Assert.Contains(cells[0].InnerCell!.Descendants<Text>(), t => t.Text == "cell-ctl");
            Assert.IsType<TableCell>(cells[1].CellChild);
            Assert.Same(cells[1].CellChild, cells[1].InnerCell);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Cells_reports_the_declared_gridSpan_and_treats_a_missing_one_as_one()
    {
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
            + "<w:tr>"
            + "<w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p><w:r><w:t>spanned</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr/><w:p><w:r><w:t>plain</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var table = FirstTable(path);
            var innerRow = TableGridNavigator.Rows(table).Single().InnerRow!;
            var cells = TableGridNavigator.Cells(innerRow);

            Assert.Equal(2, cells[0].GridSpan);
            Assert.Equal(1, cells[1].GridSpan);
            Assert.Equal(3, TableGridNavigator.GridColumnCount(table));
            Assert.Equal(3, cells.Sum(c => c.GridSpan)); // covers the declared grid exactly
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- GridColumnCount / GridEdges ----

    [Fact]
    public void GridColumnCount_counts_every_gridCol_even_one_with_no_declared_width()
    {
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr/><w:p/></w:tc><w:tc><w:tcPr/><w:p/></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var table = FirstTable(path);
            Assert.Equal(2, TableGridNavigator.GridColumnCount(table));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GridColumnCount_is_zero_when_the_table_has_no_tblGrid()
    {
        var body = "<w:tbl><w:tblPr/><w:tr><w:tc><w:tcPr/><w:p/></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var table = FirstTable(path);
            Assert.Equal(0, TableGridNavigator.GridColumnCount(table));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GridEdges_reads_gridBefore_and_gridAfter_from_the_rows_trPr()
    {
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
            + "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/></w:trPr>"
            + "<w:tc><w:tcPr/><w:p><w:r><w:t>c1</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr/><w:p><w:r><w:t>c2</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var table = FirstTable(path);
            var innerRow = TableGridNavigator.Rows(table).Single().InnerRow!;
            var (before, after) = TableGridNavigator.GridEdges(innerRow);
            var cells = TableGridNavigator.Cells(innerRow);

            Assert.Equal(1, before);
            Assert.Equal(0, after);
            Assert.Equal(3, before + after + cells.Sum(c => c.GridSpan)); // covers the declared grid exactly
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GridEdges_defaults_to_zero_when_trPr_has_no_gridBefore_or_gridAfter()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 1, cols: 1));
        try
        {
            var table = FirstTable(path);
            var innerRow = TableGridNavigator.Rows(table).Single().InnerRow!;
            var (before, after) = TableGridNavigator.GridEdges(innerRow);

            Assert.Equal(0, before);
            Assert.Equal(0, after);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- Tables() / TableAt(): table-by-index resolution ----

    [Fact]
    public void Tables_flattens_a_nested_table_into_document_order_with_its_own_index()
    {
        var outerCellWithNestedTable =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr/>"
            + SyntheticLayout.SimpleTable(rows: 1, cols: 1)
            + "<w:p/></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(outerCellWithNestedTable);
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var body = doc.MainDocumentPart!.Document!.Body!;

            var tables = TableGridNavigator.Tables(body);
            Assert.Equal(2, tables.Count);
            Assert.Same(TableGridNavigator.TableAt(body, 0, "the document body"), tables[0]);
            Assert.Same(TableGridNavigator.TableAt(body, 1, "the document body"), tables[1]);

            // The nested table (index 1) is inside the outer table's (index 0) one cell.
            var outerCell = TableGridNavigator.Rows(tables[0]).Single().InnerRow!
                .Elements<TableCell>().Single();
            Assert.Contains(outerCell.Descendants<Table>(), t => ReferenceEquals(t, tables[1]));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tables_finds_a_table_nested_inside_an_sdt_block()
    {
        var body =
            "<w:sdt><w:sdtPr><w:id w:val=\"600\"/></w:sdtPr><w:sdtContent>"
            + SyntheticLayout.SimpleTable(rows: 1, cols: 1)
            + "</w:sdtContent></w:sdt>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var docBody = doc.MainDocumentPart!.Document!.Body!;

            var tables = TableGridNavigator.Tables(docBody);
            var table = Assert.Single(tables);
            Assert.IsType<SdtBlock>(table.Parent!.Parent); // w:tbl -> w:sdtContent -> w:sdt
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableAt_returns_the_table_at_the_given_zero_based_index()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(1, 1) + SyntheticLayout.SimpleTable(1, 1));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var expected = body.Descendants<Table>().ElementAt(1);

            Assert.Same(expected, TableGridNavigator.TableAt(body, 1, "the document body"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableAt_out_of_range_throws_NotFoundException_with_TableCoordinate_target()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(1, 1));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var body = doc.MainDocumentPart!.Document!.Body!;

            var ex = Assert.Throws<NotFoundException>(() => TableGridNavigator.TableAt(body, 5, "the document body"));
            Assert.Contains("Table index 5", ex.Message);
            Assert.Contains("1 table(s)", ex.Message);
            Assert.Equal(NotFoundTarget.TableCoordinate, ex.TargetKind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static int? ReadId(SdtRow sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtId>()?.Val?.Value;
}
