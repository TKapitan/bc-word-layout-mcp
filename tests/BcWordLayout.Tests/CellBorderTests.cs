using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="TableStructureEditor.SetCellBorders"/> and its <c>set_cell_borders</c> tool surface —
/// the per-cell half of the BC table look. Real BC documents draw no table grid: their rules
/// are per-cell <c>w:tcBorders</c> (a line under the lines-table header row, a line above a totals block),
/// which <c>insert_repeater_table</c>'s <c>look='bc'</c> default only covers for the header row. Everything
/// else an author needs is this tool.
/// </summary>
public class CellBorderTests
{
    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-cellborders-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static List<ValidationErrorInfo> OpenXmlErrors(WordprocessingDocument doc) =>
        new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();

    /// <summary>Authors a fresh 2x3 plain table to draw rules on, returning its addressable table index.</summary>
    private static int NewPlainTable(string path, int rows = 2, int columns = 3)
    {
        var response = TableTools.InsertTable(path, rows, columns, locationType: "documentEnd");
        Assert.True(response.Ok, response.Error?.Message);
        return Assert.IsType<TableEditResultDto>(response.Data).TableIndex;
    }

    private static List<TableCell> RowCells(WordprocessingDocument doc, int tableIndex, int row) =>
        TableGridNavigator.Tables(doc.MainDocumentPart!.Document!.Body!)[tableIndex]
            .Elements<TableRow>()
            .ElementAt(row)
            .Elements<TableCell>()
            .ToList();

    [Fact]
    public void Tool_set_cell_borders_without_col_draws_the_rule_across_every_cell_in_the_row()
    {
        // The usual case: "put a line above the totals row" is a whole-row rule, not a per-cell chore.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var tableIndex = NewPlainTable(path);

            var response = TableTools.SetCellBorders(path, tableIndex, row: 1, edges: "top");
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);
            Assert.Equal("set_cell_borders", dto.Operation);
            Assert.Null(dto.ColumnIndex);
            Assert.Equal(3, dto.ColumnCountBefore);
            Assert.Equal(3, dto.ColumnCountAfter); // cosmetic only — the grid is untouched

            using var reopened = WordprocessingDocument.Open(path, false);
            foreach (var cell in RowCells(reopened, tableIndex, 1))
            {
                var top = cell.TableCellProperties!.TableCellBorders!.TopBorder!;
                Assert.Equal(BorderValues.Single, top.Val!.Value);
                Assert.Equal(4u, top.Size!.Value); // the BC-standard ½ pt
                Assert.Equal("auto", top.Color!.Value);
            }

            // The other row is untouched.
            Assert.All(RowCells(reopened, tableIndex, 0), c => Assert.Null(c.TableCellProperties?.TableCellBorders));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_set_cell_borders_with_col_targets_one_cell_and_keeps_its_unnamed_edges()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var tableIndex = NewPlainTable(path);

            Assert.True(TableTools.SetCellBorders(path, tableIndex, row: 0, edges: "bottom", col: 1).Ok);
            var second = TableTools.SetCellBorders(path, tableIndex, row: 0, edges: "top", col: 1);
            Assert.True(second.Ok, second.Error?.Message);
            Assert.Equal(1, Assert.IsType<TableEditResultDto>(second.Data).ColumnIndex);

            using var reopened = WordprocessingDocument.Open(path, false);
            var cells = RowCells(reopened, tableIndex, 0);

            // Both rules survive: an edge the second call did not name keeps whatever it had.
            var borders = cells[1].TableCellProperties!.TableCellBorders!;
            Assert.Equal(BorderValues.Single, borders.TopBorder!.Val!.Value);
            Assert.Equal(BorderValues.Single, borders.BottomBorder!.Val!.Value);

            Assert.Null(cells[0].TableCellProperties?.TableCellBorders);
            Assert.Null(cells[2].TableCellProperties?.TableCellBorders);
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_set_cell_borders_style_none_clears_an_edge_with_an_explicit_nil()
    {
        // An explicit nil, not a deleted element: in a table that declares its own w:tblBorders, a deleted
        // edge simply inherits the grid line back.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var inserted = TableTools.InsertTable(path, rows: 1, columns: 2, locationType: "documentEnd", withBorders: true);
            var tableIndex = Assert.IsType<TableEditResultDto>(inserted.Data).TableIndex;

            var response = TableTools.SetCellBorders(path, tableIndex, row: 0, edges: "left,right", style: "none");
            Assert.True(response.Ok, response.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            foreach (var cell in RowCells(reopened, tableIndex, 0))
            {
                var borders = cell.TableCellProperties!.TableCellBorders!;
                Assert.Equal(BorderValues.Nil, borders.LeftBorder!.Val!.Value);
                Assert.Equal(BorderValues.Nil, borders.RightBorder!.Val!.Value);
                Assert.Null(borders.TopBorder);
            }

            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_set_cell_borders_lands_after_the_cells_width_in_tcPr_schema_order()
    {
        // CT_TcPr sequences tcW before tcBorders; a freshly authored cell always carries a width, so
        // inserting the borders element at the front (the pre-B32 helper's only behavior for a non-gridSpan
        // child) would trip the pre-save validator gate exactly like merge_cells' w:gridSpan once did.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var tableIndex = NewPlainTable(path, rows: 1, columns: 2);
            Assert.True(TableTools.SetCellBorders(path, tableIndex, row: 0, edges: "all").Ok);

            using var reopened = WordprocessingDocument.Open(path, false);
            var tcPr = RowCells(reopened, tableIndex, 0)[0].TableCellProperties!;
            var order = tcPr.ChildElements.Select(e => e.LocalName).ToList();
            Assert.Equal(new[] { "tcW", "tcBorders" }, order);
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_set_cell_borders_rejects_an_unknown_edge_and_an_out_of_range_size()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var tableIndex = NewPlainTable(path);

            var badEdge = TableTools.SetCellBorders(path, tableIndex, row: 0, edges: "diagonal");
            Assert.False(badEdge.Ok);
            Assert.Equal("invalid_argument", badEdge.Error!.Code);
            Assert.Contains("top", badEdge.Error.Hint, StringComparison.OrdinalIgnoreCase);

            var noEdge = TableTools.SetCellBorders(path, tableIndex, row: 0, edges: "");
            Assert.False(noEdge.Ok);
            Assert.Equal("invalid_argument", noEdge.Error!.Code);

            var badSize = TableTools.SetCellBorders(path, tableIndex, row: 0, edges: "top", size: 500);
            Assert.False(badSize.Ok);
            Assert.Equal("invalid_argument", badSize.Error!.Code);
            Assert.Contains("eighths", badSize.Error.Hint, StringComparison.OrdinalIgnoreCase);

            var badStyle = TableTools.SetCellBorders(path, tableIndex, row: 0, edges: "top", style: "dotted");
            Assert.False(badStyle.Ok);
            Assert.Equal("invalid_argument", badStyle.Error!.Code);

            var badRow = TableTools.SetCellBorders(path, tableIndex, row: 99, edges: "top");
            Assert.False(badRow.Ok);
            Assert.Equal("not_found", badRow.Error!.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellBorders_works_on_a_vMerge_table_that_the_other_table_tools_reject()
    {
        // Cosmetic only — no span arithmetic — so the vMerge/gridBefore rejection every other table
        // operation carries would be pure lost capability here.
        var path = SyntheticLayout.Create(VerticallyMergedTable);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var result = TableStructureEditor.SetCellBorders(
                    doc, LayoutPart.Body, partName: null, tableIndex: 0, row: 0, col: null,
                    new CellBorderOptions { Bottom = true });

                Assert.Equal("set_cell_borders", result.Operation);
                Assert.Contains("all 2 cell(s) of row 0", result.Summary, StringComparison.Ordinal);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var cells = RowCells(reopened, 0, 0);
            Assert.All(cells, c => Assert.NotNull(c.TableCellProperties!.TableCellBorders!.BottomBorder));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>A two-row, two-column table whose first column is vertically merged (<c>w:vMerge</c>).</summary>
    private const string VerticallyMergedTable =
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
        + "<w:tr>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/><w:vMerge w:val=\"restart\"/></w:tcPr>"
        + "<w:p><w:r><w:t>merged</w:t></w:r></w:p></w:tc>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>"
        + "</w:tr>"
        + "<w:tr>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/><w:vMerge/></w:tcPr><w:p/></w:tc>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>"
        + "</w:tr></w:tbl>";
}
