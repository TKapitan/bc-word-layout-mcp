using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #81: <c>get_layout_info</c> reported only a cell's PHYSICAL position in its row
/// (<c>colIndex</c>), while the column tools take a GRID column, and real BC tables are
/// <c>gridSpan</c>-pervasive — so nothing in the response let an agent find the index
/// <c>remove_column</c>/<c>insert_column</c>/<c>set_column_widths</c> act on. Each cell now also reports the
/// grid column it starts at and its span, and each row its skipped leading/trailing grid columns.
/// </summary>
/// <remarks>
/// The fixture is the issue's own: table 2 of <c>StandardSalesInvoiceVatSpec.docx</c>, the stock VAT-spec
/// invoice's lines table (11 grid columns). Line Amount is physical cell 7 in both its header and data rows
/// but grid column 9: the header's VAT % cell spans grid 6–8, and in the data row Line Discount spans 5–6
/// and VAT % 7–8. Rows 0–2 end in a one-column <c>w:gridAfter</c> run.
/// </remarks>
public class GridCoordinateReportingTests
{
    private const int LinesTable = 2;
    private const int LineAmountLabelId = 1017885413; // #Nav: /Header/Line/LineAmount_Line_Lbl (header row)
    private const int LineAmountFieldId = -1829895906; // #Nav: /Header/Line/LineAmount_Line (data row)

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-gridcoord-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static TableDto LinesTableInfo(string path) =>
        Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data).Tables
            .Single(t => t.Part == "document.xml" && t.TableIndex == LinesTable);

    private static TableCellDto CellOf(TableDto table, int controlId) =>
        table.Rows.SelectMany(r => r.Cells).Single(c => c.ControlId == controlId);

    [Fact]
    public void GetLayoutInfo_reports_each_cells_grid_column_and_span_next_to_its_physical_index()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoiceVatSpec);
        try
        {
            var table = LinesTableInfo(path);
            Assert.Equal(11, table.ColumnCount);

            foreach (var id in new[] { LineAmountLabelId, LineAmountFieldId })
            {
                var lineAmount = CellOf(table, id);
                Assert.Equal(7, lineAmount.ColIndex);
                Assert.Equal(9, lineAmount.GridColumn);
                Assert.Equal(1, lineAmount.GridSpan);
            }

            var header = table.Rows[0].Cells;
            Assert.Equal((6, 6, 3), (header[6].ColIndex, header[6].GridColumn, header[6].GridSpan)); // VAT %

            var data = table.Rows[1].Cells;
            Assert.Equal((5, 5, 2), (data[5].ColIndex, data[5].GridColumn, data[5].GridSpan)); // Line Discount %
            Assert.Equal((6, 7, 2), (data[6].ColIndex, data[6].GridColumn, data[6].GridSpan)); // VAT %

            Assert.Equal(new[] { 1, 1, 1, 0 }, table.Rows.Select(r => r.GridAfter).ToArray());
            Assert.All(table.Rows, r => Assert.Equal(0, r.GridBefore));

            // The fields account for the whole grid on every row: before + Σspan + after = columnCount.
            Assert.All(table.Rows, r =>
                Assert.Equal(table.ColumnCount, r.GridBefore + r.Cells.Sum(c => c.GridSpan) + r.GridAfter));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GetLayoutInfo_starts_a_gridBefore_rows_first_cell_at_the_skipped_offset()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithSkippedGridCells());
        try
        {
            var table = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data).Tables.Single();

            Assert.Equal((0, 2), (table.Rows[1].GridBefore, table.Rows[1].GridAfter));
            Assert.Equal(new[] { 0, 1 }, table.Rows[1].Cells.Select(c => c.GridColumn).ToArray());

            Assert.Equal((1, 0), (table.Rows[2].GridBefore, table.Rows[2].GridAfter));
            Assert.Equal(new[] { 1, 2, 3 }, table.Rows[2].Cells.Select(c => c.GridColumn).ToArray());
            Assert.Equal(new[] { 0, 1, 2 }, table.Rows[2].Cells.Select(c => c.ColIndex).ToArray());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Remove_column_given_the_reported_grid_column_removes_the_column_the_agent_asked_for()
    {
        // The round trip SOLUTION-DESIGN §4 promises: read Line Amount's gridColumn, pass it to remove_column,
        // and Line Amount is what goes — not a column of the VAT % cells (what colIndex 7 removed).
        var path = CopyOfCorpus(Corpus.SalesInvoiceVatSpec);
        try
        {
            var column = CellOf(LinesTableInfo(path), LineAmountFieldId).GridColumn;

            var response = TableTools.RemoveColumn(path, LinesTable, column);
            Assert.True(response.Ok, response.Error?.Message);
            var summary = Assert.IsType<TableEditResultDto>(response.Data).Summary;
            Assert.Contains("deleted the covering cell in rows 0, 1, 2", summary);
            Assert.Contains("narrowed the covering spanned cell by one column in row 3", summary);
            Assert.DoesNotContain("No row had a cell of its own", summary);

            using var doc = WordprocessingDocument.Open(path, false);
            var ids = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtElement>()
                .Select(SdtInspector.ReadControlId).ToHashSet();
            Assert.DoesNotContain(LineAmountLabelId, ids);
            Assert.DoesNotContain(LineAmountFieldId, ids);

            // The VAT % cells kept their spans: header 3, data row 2.
            var table = TableGridNavigator.Tables(doc.MainDocumentPart.Document.Body!)[LinesTable];
            var rows = IndependentGridWalk.Rows(table);
            Assert.Equal(10, IndependentGridWalk.GridCount(table));
            Assert.Equal(3, IndependentGridWalk.Span(IndependentGridWalk.Cells(rows[0])[6]));
            Assert.Equal(2, IndependentGridWalk.Span(IndependentGridWalk.Cells(rows[1])[6]));
            Assert.All(rows, r => Assert.Equal(10, IndependentGridWalk.Coverage(r)));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Remove_column_summary_warns_when_no_row_had_a_cell_of_its_own_at_that_grid_column()
    {
        // The issue's repro: column 7 is Line Amount's PHYSICAL index. Every row covers grid column 7 with a
        // wider spanned cell, so nothing the agent meant was removed — the summary now says so.
        var path = CopyOfCorpus(Corpus.SalesInvoiceVatSpec);
        try
        {
            var response = TableTools.RemoveColumn(path, LinesTable, column: 7);
            Assert.True(response.Ok, response.Error?.Message);

            var summary = Assert.IsType<TableEditResultDto>(response.Data).Summary;
            Assert.Contains("updating 4 row(s) (narrowed the covering spanned cell by one column in rows 0, 1, 2, 3)", summary);
            Assert.Contains("No row had a cell of its own in grid column 7", summary);
            Assert.Contains("gridColumn", summary);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Remove_column_summary_names_the_rows_whose_skipped_run_shrank()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoiceVatSpec);
        try
        {
            // Grid column 10 is the 34-twip trailing column: rows 0-2 reach it through their w:gridAfter run,
            // the totals row through its last, spanned cell.
            var response = TableTools.RemoveColumn(path, LinesTable, column: 10);
            Assert.True(response.Ok, response.Error?.Message);

            var summary = Assert.IsType<TableEditResultDto>(response.Data).Summary;
            Assert.Contains("narrowed the covering spanned cell by one column in row 3", summary);
            Assert.Contains("shrank the skipped grid run in rows 0, 1, 2", summary);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(6, 7)] // VAT %, spanning grid 7-8 in the data row
    [InlineData(7, 9)] // Line Amount
    public void Set_cell_format_reports_the_grid_column_of_the_cell_it_draws_on(int physicalCol, int gridColumn)
    {
        // TableEditResultDto.ColumnIndex is documented as a GRID column; set_cell_borders, as the tool was
        // named then, echoed the physical col back instead (GitHub issue #81).
        var path = CopyOfCorpus(Corpus.SalesInvoiceVatSpec);
        try
        {
            var response = TableTools.SetCellFormat(path, LinesTable, row: 1, edges: "bottom", col: physicalCol);
            Assert.True(response.Ok, response.Error?.Message);

            var dto = Assert.IsType<TableEditResultDto>(response.Data);
            Assert.Equal(gridColumn, dto.ColumnIndex);
            Assert.Equal(gridColumn, LinesTableInfo(path).Rows[1].Cells[physicalCol].GridColumn);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
