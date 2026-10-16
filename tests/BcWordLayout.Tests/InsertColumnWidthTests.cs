using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// How <c>insert_column</c> reconciles widths (GitHub issue #83): the table keeps its total width, the new
/// column's width is taken proportionally from the existing columns, and every well-formed row is re-sliced
/// from the final grid — so a percentage-width table no longer squeezes the new column to a sliver and a
/// full-width table no longer grows past the page margin.
/// </summary>
public class InsertColumnWidthTests
{
    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-colwidths-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static Table BodyTable(WordprocessingDocument doc, int index) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().ElementAt(index);

    private static List<int> GridWidths(Table table) =>
        table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Select(g => int.Parse(g.Width!.Value!)).ToList();

    /// <summary>
    /// Every well-formed row's cells carry a dxa <c>w:tcW</c> equal to the grid columns they cover — the
    /// invariant remove_column and set_column_widths already keep, and the one issue #83 found broken.
    /// </summary>
    private static void AssertEveryRowIsSlicedFromTheGrid(Table table)
    {
        var grid = GridWidths(table);
        foreach (var row in TableGridNavigator.Rows(table).Where(r => r.InnerRow is not null))
        {
            var coverage = TableGridNavigator.Coverage(row.InnerRow!);
            Assert.Equal(grid.Count, coverage.Total);

            var cursor = coverage.Before;
            foreach (var cell in TableGridNavigator.Cells(row.InnerRow!))
            {
                var tcW = cell.InnerCell!.GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!;
                Assert.Equal(TableWidthUnitValues.Dxa, tcW.Type!.Value);
                Assert.Equal(grid.Skip(cursor).Take(cell.GridSpan).Sum(), int.Parse(tcW.Width!.Value!));
                cursor += cell.GridSpan;
            }
        }
    }

    [Fact]
    public void A_percentage_width_table_gives_the_new_column_an_equal_share_instead_of_a_sliver()
    {
        // The issue's first repro: the quote's info block is 5 cells of w:tcW 1000 pct (20% each). The new
        // dxa cell used to sit beside cells still claiming 100% of the table between them.
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var totalBefore = GridWidths(BodyTable(doc, 1)).Sum();

            TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, 1, atColumn: null, new InsertColumnOptions { Mode = InsertColumnMode.PlainText });

            var table = BodyTable(doc, 1);
            var grid = GridWidths(table);
            Assert.Equal(totalBefore, grid.Sum());
            Assert.Equal((int)Math.Round(totalBefore / 6.0), grid[^1]); // an equal share of the kept total
            Assert.DoesNotContain(
                table.Descendants<TableCellWidth>(), w => w.Type?.Value == TableWidthUnitValues.Pct);
            AssertEveryRowIsSlicedFromTheGrid(table);

            // The table itself stays a 100%-wide table.
            var tblW = table.GetFirstChild<TableProperties>()!.TableWidth!;
            Assert.Equal(TableWidthUnitValues.Pct, tblW.Type!.Value);
            Assert.Equal("5000", tblW.Width!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_full_width_dxa_table_keeps_its_total_width_and_the_new_column_its_requested_width()
    {
        // The issue's second repro: tblW went 10262 -> 11462, clipping the new column at the page edge and
        // pushing the right-anchored grand total off the page.
        var path = CopyOfCorpus(Corpus.SalesInvoiceVatSpec);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var gridBefore = GridWidths(BodyTable(doc, 2));

            var result = TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, 2, atColumn: null,
                new InsertColumnOptions { Mode = InsertColumnMode.Field, DataPath = "/Header/Line/ShipmentDate_Line", Width = 1200 });

            var table = BodyTable(doc, 2);
            var grid = GridWidths(table);
            Assert.Equal("10262", table.GetFirstChild<TableProperties>()!.TableWidth!.Width!.Value);
            Assert.Equal(gridBefore.Sum(), grid.Sum());

            // The content rows end in a 34-twip skipped column; their new cells land just before it, so the
            // new w:gridCol goes there too (grid 10) and the skipped column stays last.
            Assert.Equal(10, result.ColumnIndex);
            Assert.Equal(1200, grid[10]);
            Assert.True(grid[11] < gridBefore[10] + 1, "the trailing filler column must stay the narrow last one");
            Assert.Contains("ahead of the rows' trailing skipped grid column", result.Summary, StringComparison.Ordinal);

            var lineRow = TableGridNavigator.Rows(table)[1].InnerRow!;
            Assert.Equal(1, TableGridNavigator.Coverage(lineRow).After);
            var newCell = TableGridNavigator.Cells(lineRow)[^1].InnerCell!;
            Assert.Equal("1200", newCell.GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!.Width!.Value);

            AssertEveryRowIsSlicedFromTheGrid(table);
            Assert.Empty(TableGridConsistencyGuard.Find(doc));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void The_default_width_is_an_equal_share_of_the_kept_total()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3)); // 3 x 2000 twips
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var result = TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, 0, atColumn: 1, new InsertColumnOptions { Mode = InsertColumnMode.PlainText });

            Assert.Equal(new[] { 1500, 1500, 1500, 1500 }, GridWidths(BodyTable(doc, 0)));
            Assert.Contains("table width kept at 6000 twips", result.Summary, StringComparison.Ordinal);
            AssertEveryRowIsSlicedFromTheGrid(BodyTable(doc, 0));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-5)]
    [InlineData(6000)] // the whole table: nothing would be left for the existing columns
    public void A_width_that_leaves_no_room_for_the_existing_columns_is_refused(int width)
    {
        var path = SyntheticLayout.Create(SyntheticLayout.SimpleTable(rows: 2, cols: 3));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => TableStructureEditor.InsertColumn(
                doc, LayoutPart.Body, null, 0, atColumn: null,
                new InsertColumnOptions { Mode = InsertColumnMode.PlainText, Width = width }));

            Assert.Equal("width", ex.ParamName);
            Assert.Equal(3, TableGridNavigator.GridColumnCount(BodyTable(doc, 0))); // untouched
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_insert_column_on_the_VatSpec_lines_table_keeps_tblW()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoiceVatSpec);
        try
        {
            var response = TableTools.InsertColumn(
                path, tableIndex: 2, mode: "field", dataPath: "/Header/Line/ShipmentDate_Line", width: 1200);
            Assert.True(response.Ok, response.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.Equal("10262", BodyTable(reopened, 2).GetFirstChild<TableProperties>()!.TableWidth!.Width!.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
