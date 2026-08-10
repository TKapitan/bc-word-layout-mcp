using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="LayoutEditor.InsertPlainTable"/> and its <c>insert_table</c> tool surface — the
/// building block for authoring the NON-repeating sections of a layout from scratch (address columns,
/// header-info grids, totals blocks), born from the 2026-07-31 from-scratch authoring exercise where no
/// tool could produce them. The key contract: the result's TableIndex must be immediately addressable by
/// set_cell_text / insert_field (locationType 'tableCell').
/// </summary>
public class PlainTableTests
{
    private static string CopyOfCorpus(string corpusFile)
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"bcwl-plaintable-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static List<ValidationErrorInfo> OpenXmlErrors(WordprocessingDocument doc) =>
        new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();

    [Fact]
    public void Tool_insert_table_appends_an_empty_grid_the_follow_up_cell_tools_can_address()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertTable(path, rows: 2, columns: 3, locationType: "documentEnd",
                columnWidths: "3000,4000,3206", columnAlignments: "left,center,right");
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);
            Assert.Equal("insert_table", dto.Operation);
            Assert.Equal(0, dto.ColumnCountBefore);
            Assert.Equal(3, dto.ColumnCountAfter);
            Assert.Equal(2, dto.RowsAffected);
            Assert.True(dto.QuickValidation.Passed, "fresh plain table must not introduce findings");

            // The returned index must address the new table for the follow-up cell tools.
            var setResponse = EditTools.SetCellText(path, dto.TableIndex, 0, 0, "Caption");
            Assert.True(setResponse.Ok, setResponse.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var widths = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Select(g => g.Width!.Value).ToList();
            Assert.Equal(new[] { "3000", "4000", "3206" }, widths);
            Assert.Null(table.GetFirstChild<TableProperties>()!.GetFirstChild<TableBorders>()); // borderless default

            var firstRowCells = table.Elements<TableRow>().First().Elements<TableCell>().ToList();
            Assert.Equal(3, firstRowCells.Count);
            Assert.Equal("Caption", firstRowCells[0].InnerText);
            // Per-column alignment seeds each cell's paragraph justification.
            Assert.Equal(
                JustificationValues.Center,
                firstRowCells[1].GetFirstChild<Paragraph>()!.ParagraphProperties!.Justification!.Val!.Value);
            Assert.Equal(
                JustificationValues.Right,
                firstRowCells[2].GetFirstChild<Paragraph>()!.ParagraphProperties!.Justification!.Val!.Value);

            Assert.Empty(TableGridConsistencyGuard.Find(reopened));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_insert_table_with_borders_draws_the_explicit_single_line_grid()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertTable(path, rows: 1, columns: 2, locationType: "documentEnd",
                withBorders: true);
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var borders = table.GetFirstChild<TableProperties>()!.GetFirstChild<TableBorders>();
            Assert.NotNull(borders?.InsideHorizontalBorder);
            // Default widths: an even split of the full content width.
            Assert.Equal(
                new[] { "5103", "5103" },
                table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Select(g => g.Width!.Value));
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("left,right", 3, "columnAlignments")]     // count mismatch (3 columns, 2 alignments)
    [InlineData("left,middle,right", 3, "columnAlignments")] // unknown alignment name
    public void Tool_insert_table_rejects_bad_alignment_input(string alignments, int columns, string expectedParam)
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertTable(path, rows: 1, columns: columns,
                locationType: "documentEnd", columnAlignments: alignments);
            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.Contains(expectedParam, response.Error.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tool_insert_table_rejects_header_and_footer_parts()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertTable(path, rows: 1, columns: 1,
                locationType: "documentEnd", layoutPart: "footer");
            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Inserted_control_takes_optional_bold_and_size_on_its_runs_and_sdtPr()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var tableResponse = TableTools.InsertTable(path, rows: 1, columns: 2, locationType: "documentEnd");
            var dto = Assert.IsType<TableEditResultDto>(tableResponse.Data);

            var response = EditTools.InsertLabel(
                path, "/Header/ExternalDocumentNo_Lbl", "tableCell",
                tableIndex: dto.TableIndex, row: 0, col: 0, bold: true, fontSizePoints: 8);
            Assert.True(response.Ok, response.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var sdt = table.Descendants<SdtRun>().Single();
            var runProps = sdt.Descendants<Run>().Single().RunProperties!;
            Assert.NotNull(runProps.Bold);
            Assert.Equal("16", runProps.FontSize!.Val!.Value); // 8pt = 16 half-points
            // sdtPr rPr styles the placeholder/future content; corpus order puts it first.
            var sdtPrRunProps = sdt.SdtProperties!.GetFirstChild<RunProperties>();
            Assert.NotNull(sdtPrRunProps?.Bold);
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Inserted_control_rejects_the_alignment_knob_with_guidance()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var ex = Assert.Throws<ArgumentException>(() => LayoutEditor.InsertField(
                doc, "/Header/ExternalDocumentNo",
                new Location { Type = LocationKind.DocumentEnd },
                new BcWordLayout.Domain.Models.CellTextFormat { Alignment = "right" }));
            Assert.Contains("columnAlignments", ex.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Stacked_block_tables_get_a_separator_paragraph_so_Word_never_merges_them()
    {
        // Word renders two adjacent w:tbl siblings as ONE merged table; authoring in Word always keeps a
        // paragraph between them. Stacking blocks (address grid straight after a title table, a repeater
        // straight after an info grid) must therefore never leave two tables touching.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            Assert.True(TableTools.InsertTable(path, rows: 1, columns: 2, locationType: "documentEnd").Ok);
            Assert.True(TableTools.InsertTable(path, rows: 1, columns: 3, locationType: "documentEnd").Ok);
            Assert.True(TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line,Description_Line", "documentEnd").Ok);

            using var reopened = WordprocessingDocument.Open(path, false);
            var body = reopened.MainDocumentPart!.Document!.Body!;
            var adjacentPairs = body.ChildElements.Zip(body.ChildElements.Skip(1))
                .Count(pair => pair.First is Table && pair.Second is Table);
            Assert.Equal(0, adjacentPairs);
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Fresh_plain_table_cell_takes_a_bound_field_via_tableCell_location()
    {
        // The authoring loop this tool exists for: insert_table -> insert_field into one of its cells.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var tableResponse = TableTools.InsertTable(path, rows: 1, columns: 2, locationType: "documentEnd");
            var dto = Assert.IsType<TableEditResultDto>(tableResponse.Data);

            var fieldResponse = EditTools.InsertField(
                path, "/Header/ExternalDocumentNo", "tableCell", tableIndex: dto.TableIndex, row: 0, col: 1);
            Assert.True(fieldResponse.Ok, fieldResponse.Error?.Message);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var cell = table.Elements<TableRow>().First().Elements<TableCell>().ElementAt(1);
            Assert.Single(cell.Descendants<SdtRun>());
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Fresh_plain_table_pins_its_grid_with_tblW_and_a_fixed_tblLayout()
    {
        // Without w:tblLayout the OOXML default is autofit, so Word recomputes every column from cell
        // content and the columnWidths passed here hold only until someone opens the file - one plain Word
        // save rewrote an untouched 2800,3000 grid to 3177,3066 (GitHub issue #52). Every corpus table
        // declares both elements.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertTable(
                path, rows: 2, columns: 3, locationType: "documentEnd", columnWidths: "3000,4000,3206");
            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var tblPr = table.GetFirstChild<TableProperties>()!;

            Assert.Equal(TableLayoutValues.Fixed, tblPr.TableLayout!.Type!.Value);
            Assert.Equal(TableWidthUnitValues.Dxa, tblPr.TableWidth!.Type!.Value);

            // The declared total is the grid's own sum, so the two can never disagree.
            var gridSum = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>()
                .Sum(c => int.Parse(c.Width!.Value!, System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(3000 + 4000 + 3206, gridSum);
            Assert.Equal(gridSum.ToString(System.Globalization.CultureInfo.InvariantCulture), tblPr.TableWidth.Width!.Value);

            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Bordered_plain_table_orders_tblPr_children_as_the_schema_requires()
    {
        // CT_TblPrBase is a SEQUENCE: w:tblW precedes w:tblBorders, which precedes w:tblLayout. Appending
        // the new elements instead of assigning the SDK's typed properties would emit them out of order -
        // valid-looking C#, invalid OOXML - so this asserts the order explicitly rather than relying on
        // OpenXmlValidator alone to notice.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertTable(
                path, rows: 1, columns: 2, locationType: "documentEnd", withBorders: true);
            var dto = Assert.IsType<TableEditResultDto>(response.Data);

            using var reopened = WordprocessingDocument.Open(path, false);
            var table = TableGridNavigator.Tables(reopened.MainDocumentPart!.Document!.Body!)[dto.TableIndex];
            var names = table.GetFirstChild<TableProperties>()!.ChildElements
                .Select(e => e.LocalName).ToList();

            Assert.Equal(new[] { "tblW", "tblBorders", "tblLayout" }, names);
            Assert.Empty(OpenXmlErrors(reopened));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
