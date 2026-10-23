using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #85: a cell wrapped in TWO cell-level content controls — <c>w:tr &gt; w:sdt &gt; w:sdt &gt;
/// w:tc</c> — was invisible to every table tool, because <see cref="TableGridNavigator"/> descended one
/// <c>w:sdt</c> level only and so found no <c>w:tc</c> at all. The shape is the Business Central Word add-in's
/// conditional control placed AROUND a bound field's own cell-level control: in
/// <c>SalespersonCommission.docx</c> (report 115) a "Hide Table Row If Empty" wrapper around the bound
/// <c>SalespersonPurchaser_Code</c> cell, and in <c>SubcontractorDispatchList.docx</c> a "Hide Table If
/// Empty" one around a cell spanning four grid columns.
/// </summary>
/// <remarks>
/// Every assertion about the OUTPUT walks the cells independently of the navigator
/// (<see cref="IndependentGridWalk"/>), so the tests cannot share the blind spot they are about.
/// </remarks>
public class NestedCellControlTests
{
    // SalespersonCommission, table 0, row 2 (the /Salesperson_Purchaser item's first row), cell 0.
    private const int TableIndex = 0;
    private const int Row = 2;
    private const int WrapperId = -392738406; // "Hide Table Row If Empty", #BC:HideTableRowIfEmpty
    private const int FieldId = -1216735289;  // #Nav: /Salesperson_Purchaser/Cust_Ledger_Entry/SalespersonPurchaser_Code

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-nestedcell-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    /// <summary>The <c>w:tr</c> holding <paramref name="controlId"/>'s cell-level control.</summary>
    private static TableRow RowOf(WordprocessingDocument doc, int controlId) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<SdtCell>()
            .Single(s => SdtInspector.ReadControlId(s) == controlId)
            .Ancestors<TableRow>().First();

    [Fact]
    public void Cells_descends_both_cell_level_controls_to_the_inner_w_tc_and_records_every_level()
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.SalespersonCommission), false);
        var row = RowOf(doc, WrapperId);

        var cells = TableGridNavigator.Cells(row);
        var expected = IndependentGridWalk.Cells(row);

        Assert.Equal(expected.Count, cells.Count);
        Assert.Equal(9, cells.Count);
        Assert.Same(expected[0], cells[0].InnerCell);
        Assert.Equal(
            new int?[] { WrapperId, FieldId },
            cells[0].Wrappers.Select(w => SdtInspector.ReadControlId(w)).ToArray());
        Assert.Same(cells[0].CellChild, cells[0].Wrappers[0]);
    }

    [Fact]
    public void Cells_reads_the_span_of_the_wrapped_inner_cell()
    {
        // SubcontractorDispatchList: the "Hide Table If Empty" wrapper sits around a gridSpan-4 cell. Read as
        // span 1, the row covered 7 of 10 grid columns.
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.SubcontractorDispatchList), false);
        var wrapper = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtCell>()
            .Single(s => SdtInspector.ReadAlias(s) == "Hide Table If Empty");
        var row = wrapper.Ancestors<TableRow>().First();

        var cells = TableGridNavigator.Cells(row);

        var slot = Assert.Single(cells, c => ReferenceEquals(c.CellChild, wrapper));
        Assert.Equal(4, slot.GridSpan);
        Assert.Equal(2, slot.Wrappers.Count);
        Assert.Equal(IndependentGridWalk.Cells(row).Sum(IndependentGridWalk.Span), cells.Sum(c => c.GridSpan));
        Assert.Equal(10, cells.Sum(c => c.GridSpan));
    }

    [Fact]
    public void GetLayoutInfo_reports_the_bound_field_as_the_cells_control_and_gives_every_level_its_coordinate()
    {
        var path = CopyOfCorpus(Corpus.SalespersonCommission);
        try
        {
            var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);

            var cell = info.Tables.Single(t => t.Part == "document.xml" && t.TableIndex == TableIndex)
                .Rows[Row].Cells[0];
            Assert.True(cell.IsControlCell);
            Assert.Equal(FieldId, cell.ControlId);
            Assert.Equal("Field", cell.ControlKind);
            Assert.Equal("SalespersonPurchaser_Code", cell.Text);

            foreach (var id in new[] { WrapperId, FieldId })
            {
                var control = Assert.Single(info.Controls, c => c.SdtId == id);
                Assert.Equal((TableIndex, Row, 0), (control.TableIndex, control.RowIndex, control.ColIndex));
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellFormat_draws_the_rule_on_every_w_tc_of_the_row_leaving_no_gap_under_the_wrapped_cell()
    {
        var path = CopyOfCorpus(Corpus.SalespersonCommission);
        try
        {
            var response = TableTools.SetCellFormat(path, TableIndex, Row, edges: "top");
            Assert.True(response.Ok, response.Error?.Message);
            Assert.Contains("all 9 cell(s)", Assert.IsType<TableEditResultDto>(response.Data).Summary);

            using var reopened = WordprocessingDocument.Open(path, false);
            var cells = IndependentGridWalk.Cells(RowOf(reopened, WrapperId));
            Assert.Equal(9, cells.Count);
            Assert.All(cells, c =>
                Assert.Equal(BorderValues.Single, c.TableCellProperties!.TableCellBorders!.TopBorder!.Val!.Value));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetColumnWidths_resizes_the_wrapped_cell_to_its_grid_column()
    {
        var path = CopyOfCorpus(Corpus.SalespersonCommission);
        try
        {
            var widths = new[] { 1100, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000 };
            var response = TableTools.SetColumnWidths(path, TableIndex, string.Join(",", widths));
            Assert.True(response.Ok, response.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            var wrapped = IndependentGridWalk.Cells(RowOf(reopened, WrapperId))[0];
            Assert.Equal("1100", wrapped.TableCellProperties!.TableCellWidth!.Width!.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Text_edits_addressed_at_the_wrapped_cell_are_refused_as_a_bound_control_not_reported_out_of_range()
    {
        // Before the fix both calls returned not_found with an "index out of range" hint for a cell
        // get_layout_info itself listed. The cell holds a bound field, so the honest answer is the same
        // invalid_argument refusal a single-level bound cell gets.
        var path = CopyOfCorpus(Corpus.SalespersonCommission);
        try
        {
            var before = File.ReadAllBytes(path);

            var setText = EditTools.SetCellText(path, TableIndex, Row, col: 0, text: "x");
            Assert.False(setText.Ok);
            Assert.Equal("invalid_argument", setText.Error!.Code);

            var insertText = EditTools.InsertText(path, "x", "tableCell", tableIndex: TableIndex, row: Row, col: 0);
            Assert.False(insertText.Ok);
            Assert.Equal("invalid_argument", insertText.Error!.Code);
            Assert.Contains("data-bound", insertText.Error.Message);

            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(WrapperId)]
    [InlineData(FieldId)]
    public void AfterControl_on_either_level_resolves_to_the_inner_cell_and_is_refused_as_inside_the_bound_field(int controlId)
    {
        // Was not_found: the outer wrapper's content holds no w:tc of its own ("no table cell to insert into"),
        // and the inner field's parent is the wrapper's content, not a w:tr ("cannot safely host a sibling").
        var path = CopyOfCorpus(Corpus.SalespersonCommission);
        try
        {
            var response = EditTools.InsertText(path, "x", "afterControl", controlId: controlId);

            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.Contains($"(id {FieldId})", response.Error.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
