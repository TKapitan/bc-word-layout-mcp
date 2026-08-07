using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers GitHub issue #28: static (non-repeating) rows inserted into an existing table —
/// <see cref="SdtFactory.BuildStaticRow"/>, <see cref="TableStructureEditor.InsertStaticRow"/>, and the
/// <c>insert_table_row</c> MCP tool. The target shape is the stock totals block INSIDE the line-items
/// table (StandardSalesQuote/StandardPurchaseOrder/StandardSalesInvoiceVatSpec all end their lines table
/// with right-anchored trailing rows), exercised here against StandardSalesQuote's REAL line-items table
/// — 8 grid columns, two repeating sections and pre-existing spacer/total rows — using the two Totals
/// columns the stock layout leaves unbound (<c>TotalExcludingVATText</c>, <c>TotalSubTotal</c>) so the
/// inserted controls are unambiguously distinguishable from the stock ones.
/// </summary>
[Collection("label-convention-seam")]
public class StaticTableRowTests
{
    /// <summary>StandardSalesQuote's line-items table (0-based, document order): tables 0/1 are the address/info grids.</summary>
    private const int QuoteLinesTable = 2;

    private const string CaptionPath = "/Header/Totals/TotalExcludingVATText";
    private const string AmountPath = "/Header/Totals/TotalSubTotal";

    /// <summary>4 unit spacers + a 3-column caption + the amount cell = the 8 grid columns of table 2.</summary>
    private static List<RepeaterRowCell> TotalsRowCells() =>
    [
        new RepeaterRowCell { Span = 1 },
        new RepeaterRowCell { Span = 1 },
        new RepeaterRowCell { Span = 1 },
        new RepeaterRowCell { Span = 1 },
        new RepeaterRowCell { Span = 3, Columns = [CaptionPath] },
        new RepeaterRowCell { Span = 1, Columns = [AmountPath], Alignment = "right" },
    ];

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-staticrow-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    // ---- TableStructureEditor.InsertStaticRow: the on-disk OOXML ----

    [Fact]
    public void Appended_totals_row_is_a_direct_static_child_row_with_corpus_shape_cells()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            TableEditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                result = TableStructureEditor.InsertStaticRow(
                    doc, LayoutPart.Body, null, QuoteLinesTable, atRow: null, TotalsRowCells(),
                    new CellTextFormat { Bold = true });

                // The edit introduces no structural error and no grid desync on this heavily spanned table.
                Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
                Assert.Empty(TableGridConsistencyGuard.Find(doc));

                doc.MainDocumentPart!.Document!.Save();
            }

            Assert.Equal("insert_table_row", result.Operation);
            Assert.Equal(QuoteLinesTable, result.TableIndex);
            Assert.Equal(1, result.RowsAffected);
            Assert.Equal(result.ColumnCountBefore, result.ColumnCountAfter);
            Assert.Contains("renders exactly once", result.Summary, StringComparison.Ordinal);

            using var reopened = WordprocessingDocument.Open(path, false);
            var body = reopened.MainDocumentPart!.Document!.Body!;
            var table = body.Descendants<Table>().ElementAt(QuoteLinesTable);

            // The new row is the table's LAST direct child — a bare w:tr, not inside any repeating section.
            var newRow = Assert.IsType<TableRow>(table.LastChild);
            Assert.Empty(newRow.Ancestors<SdtElement>());

            // Cell shape: 4 unit spacers, a gridSpan=3 caption, a single amount cell; widths are the sums
            // of the grid columns each covers (table 2's grid: 966,2989,897,818,1393,734,708,1701).
            var cells = newRow.Elements<TableCell>().ToList();
            Assert.Equal(6, cells.Count);
            Assert.Equal(3, cells[4].GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value);
            Assert.Equal("2835", cells[4].GetFirstChild<TableCellProperties>()?.GetFirstChild<TableCellWidth>()?.Width?.Value);
            Assert.Equal("1701", cells[5].GetFirstChild<TableCellProperties>()?.GetFirstChild<TableCellWidth>()?.Width?.Value);

            // Both bound controls carry fully-indexed xpaths (the corpus binding shape), the amount cell is
            // right-aligned, and the bold knob styled the control runs.
            var caption = Assert.IsType<SdtRun>(cells[4].Descendants<SdtRun>().Single());
            Assert.Equal(
                "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Totals[1]/ns0:TotalExcludingVATText[1]",
                SdtInspector.ReadXPath(caption));

            var amount = Assert.IsType<SdtRun>(cells[5].Descendants<SdtRun>().Single());
            Assert.Equal(
                "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Totals[1]/ns0:TotalSubTotal[1]",
                SdtInspector.ReadXPath(amount));
            Assert.Equal(
                JustificationValues.Right,
                cells[5].Descendants<Paragraph>().First().ParagraphProperties?.Justification?.Val?.Value);
            Assert.All(amount.Descendants<Run>(), r => Assert.NotNull(r.RunProperties?.Bold));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Interior_atRow_inserts_before_the_addressed_row_slot()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            int slotCountBefore;
            using (var doc = WordprocessingDocument.Open(path, false))
            {
                var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().ElementAt(QuoteLinesTable);
                slotCountBefore = TableGridNavigator.Rows(table).Count;
            }

            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // An all-spacer row (the corpus's own spacer shape) inserted mid-table, before slot 2 —
                // which in this table is the /Header/Line repeating section.
                TableStructureEditor.InsertStaticRow(
                    doc, LayoutPart.Body, null, QuoteLinesTable, atRow: 2,
                    [new RepeaterRowCell { Span = 8 }]);
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            var reopenedTable = reopened.MainDocumentPart!.Document!.Body!.Descendants<Table>().ElementAt(QuoteLinesTable);
            var slots = TableGridNavigator.Rows(reopenedTable);
            Assert.Equal(slotCountBefore + 1, slots.Count);

            // Slot 2 is the new bare row (one cell spanning all 8 columns); slot 3 is the repeater it displaced.
            Assert.False(slots[2].IsControlRow);
            var spacerCell = Assert.Single(slots[2].InnerRow!.Elements<TableCell>());
            Assert.Equal(8, spacerCell.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value);
            Assert.True(slots[3].IsControlRow);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- the row must stay STATIC through a real merge (the issue's whole point) ----

    [Fact]
    public void Merged_preview_renders_the_static_row_exactly_once_while_data_rows_multiply()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        var merged = Path.Combine(Path.GetTempPath(), $"bcwl-staticrow-merged-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                TableStructureEditor.InsertStaticRow(
                    doc, LayoutPart.Body, null, QuoteLinesTable, atRow: null, TotalsRowCells());
                doc.MainDocumentPart!.Document!.Save();
            }

            var result = MergeEngine.Merge(path, merged);
            Assert.True(result.Stats.RepeatersExpanded > 0);

            using var mergedDoc = WordprocessingDocument.Open(merged, false);
            var mergedBody = mergedDoc.MainDocumentPart!.Document!.Body!;

            // The inserted controls appear exactly ONCE in the whole merged document — a repeating section
            // would have cloned them once per data row.
            Assert.Single(mergedBody.Descendants<SdtElement>()
                .Where(s => SdtInspector.ReadXPath(s)?.Contains("TotalSubTotal", StringComparison.Ordinal) == true));

            // While the line rows genuinely multiplied (default sample data is 3 rows per repeater).
            var lineControls = mergedBody.Descendants<SdtElement>()
                .Count(s => SdtInspector.ReadXPath(s)?.Contains("ns0:ItemNo_Line", StringComparison.Ordinal) == true);
            Assert.True(lineControls > 1, $"expected the /Header/Line rows to multiply, found {lineControls}");
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(merged))
            {
                File.Delete(merged);
            }
        }
    }

    // ---- refusals (file untouched) ----

    [Fact]
    public void Spans_not_covering_the_grid_are_refused()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => TableStructureEditor.InsertStaticRow(
                doc, LayoutPart.Body, null, QuoteLinesTable, atRow: null,
                [new RepeaterRowCell { Span = 3 }]));
            Assert.Contains("spans sum to 3", ex.Message, StringComparison.Ordinal);
            Assert.Contains("8 grid column(s)", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Out_of_range_atRow_is_refused_naming_the_valid_range()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => TableStructureEditor.InsertStaticRow(
                doc, LayoutPart.Body, null, QuoteLinesTable, atRow: 99,
                [new RepeaterRowCell { Span = 8 }]));
            Assert.Contains("atRow 99 is out of range", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_row_level_alignment_is_refused_even_for_an_all_spacer_row()
    {
        // Alignment is per-cell (RepeaterRowCell.Alignment); a row-level one would be silently swallowed
        // for a spacer row (no controls to style), so it must be refused up front instead.
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => TableStructureEditor.InsertStaticRow(
                doc, LayoutPart.Body, null, QuoteLinesTable, atRow: null,
                [new RepeaterRowCell { Span = 8 }],
                new CellTextFormat { Alignment = "right" }));
            Assert.Contains("per-cell alignments", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_vMerge_table_is_refused_like_every_other_structure_op()
    {
        var body = "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr/><w:p/></w:tc></w:tr>"
            + "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>"
            + "<w:tc><w:tcPr/><w:p/></w:tc></w:tr>"
            + "</w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => TableStructureEditor.InsertStaticRow(
                doc, LayoutPart.Body, null, 0, atRow: null, [new RepeaterRowCell { Span = 2 }]));
            Assert.Contains("vMerge", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void An_unresolvable_dataset_path_is_refused()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => TableStructureEditor.InsertStaticRow(
                doc, LayoutPart.Body, null, QuoteLinesTable, atRow: null,
                [new RepeaterRowCell { Span = 8, Columns = ["/Header/Totals/NoSuchColumn"] }]));
            Assert.Contains("does not resolve", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- the insert_table_row MCP tool ----

    [Fact]
    public void InsertTableRow_tool_appends_the_stock_totals_shape_end_to_end()
    {
        var path = CopyOfCorpus(Corpus.StandardSalesQuote);
        try
        {
            // StandardSalesQuote ships with two Microsoft-shipped defects of its own (the orphaned
            // CompanyABNNumber binding), so "passes validation" is not the honest bar — "introduces no
            // NEW finding versus the pristine layout" is (the same rule the e2e scenarios apply).
            var baseline = Assert.IsType<ValidationResultDto>(ReadTools.ValidateLayout(path, "quick").Data);

            var response = TableTools.InsertTableRow(
                path,
                QuoteLinesTable,
                cells: $"-,-,-,-,3:{CaptionPath},{AmountPath}",
                alignments: "-,-,-,-,-,right",
                bold: true);

            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);
            Assert.Equal("insert_table_row", dto.Operation);
            Assert.Equal(QuoteLinesTable, dto.TableIndex);
            Assert.Equal(1, dto.RowsAffected);
            Assert.Equal(baseline.ErrorCount, dto.QuickValidation.ErrorCount);
            Assert.Equal(baseline.WarningCount, dto.QuickValidation.WarningCount);

            // The bound cells resolve against the real schema — provable by the deep (dry-run merge)
            // level: no finding may mention either inserted path.
            var full = ReadTools.ValidateLayout(path, "full");
            Assert.True(full.Ok, full.Error?.Message);
            var fullDto = Assert.IsType<ValidationResultDto>(full.Data);
            Assert.DoesNotContain(fullDto.Findings, f =>
                (f.Message + f.Location).Contains("TotalExcludingVATText", StringComparison.Ordinal)
                || (f.Message + f.Location).Contains("TotalSubTotal", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertTableRow_tool_rejects_a_malformed_span_prefix_with_invalid_argument()
    {
        var response = TableTools.InsertTableRow(
            Corpus.Path(Corpus.StandardSalesQuote), QuoteLinesTable, cells: "x:oops");

        Assert.False(response.Ok);
        Assert.Equal("invalid_argument", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }
}
